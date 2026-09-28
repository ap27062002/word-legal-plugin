"use client";

import Script from "next/script";
import { useCallback, useState } from "react";

interface Finding {
  quote: string;
  issue: string;
  severity: "low" | "medium" | "high";
  suggestion: string;
  rulebookCitation: string;
}

type FindingStatus = "pending" | "applied" | "not_found" | "error";

export default function TaskPane() {
  const [officeReady, setOfficeReady] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [findings, setFindings] = useState<Finding[]>([]);
  const [statuses, setStatuses] = useState<Record<number, FindingStatus>>({});

  const handleOfficeScriptLoad = useCallback(() => {
    Office.onReady(() => setOfficeReady(true));
  }, []);

  const runReview = useCallback(async () => {
    setLoading(true);
    setError(null);
    setFindings([]);
    setStatuses({});

    try {
      const documentText = await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        return body.text;
      });

      const response = await fetch("/api/review", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ documentText }),
      });

      if (!response.ok) {
        throw new Error(`Review request failed (${response.status})`);
      }

      const data = (await response.json()) as { findings: Finding[] };
      setFindings(data.findings);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Something went wrong");
    } finally {
      setLoading(false);
    }
  }, []);

  const applyFinding = useCallback(async (finding: Finding, index: number) => {
    try {
      const found = await Word.run(async (context) => {
        const results = context.document.body.search(finding.quote, { matchCase: true });
        results.load("items");
        await context.sync();

        if (results.items.length === 0) {
          return false;
        }

        const range = results.items[0];
        range.insertComment(
          `[Rulebook] ${finding.issue}\n\nSuggested revision: ${finding.suggestion}\n\nBasis: ${finding.rulebookCitation}`
        );
        await context.sync();
        return true;
      });

      setStatuses((prev) => ({ ...prev, [index]: found ? "applied" : "not_found" }));
    } catch {
      setStatuses((prev) => ({ ...prev, [index]: "error" }));
    }
  }, []);

  return (
    <>
      <Script
        src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"
        onLoad={handleOfficeScriptLoad}
        strategy="afterInteractive"
      />
      <main style={{ fontFamily: "Segoe UI, sans-serif", padding: 12, fontSize: 13 }}>
        <h2 style={{ fontSize: 16 }}>Rulebook Review</h2>

        <button onClick={runReview} disabled={!officeReady || loading}>
          {loading ? "Reviewing..." : "Review Document Against Rulebook"}
        </button>

        {!officeReady && <p>Loading Office...</p>}
        {error && <p style={{ color: "crimson" }}>{error}</p>}

        {findings.length === 0 && !loading && officeReady && (
          <p style={{ color: "#666" }}>No review run yet, or no issues found.</p>
        )}

        <ul style={{ listStyle: "none", padding: 0 }}>
          {findings.map((finding, index) => {
            const status = statuses[index] ?? "pending";
            return (
              <li
                key={index}
                style={{
                  border: "1px solid #ddd",
                  borderRadius: 6,
                  padding: 10,
                  marginBottom: 8,
                }}
              >
                <div style={{ fontWeight: 600 }}>
                  [{finding.severity.toUpperCase()}] {finding.issue}
                </div>
                <blockquote style={{ margin: "6px 0", color: "#555", fontStyle: "italic" }}>
                  &ldquo;{finding.quote}&rdquo;
                </blockquote>
                <div>
                  <strong>Suggestion:</strong> {finding.suggestion}
                </div>
                <div style={{ color: "#888", fontSize: 12 }}>{finding.rulebookCitation}</div>

                <button
                  onClick={() => applyFinding(finding, index)}
                  disabled={status === "applied"}
                  style={{ marginTop: 8 }}
                >
                  {status === "applied"
                    ? "Comment added"
                    : status === "not_found"
                      ? "Text not found — retry"
                      : status === "error"
                        ? "Failed — retry"
                        : "Insert as comment"}
                </button>
              </li>
            );
          })}
        </ul>
      </main>
    </>
  );
}
