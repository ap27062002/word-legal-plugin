"use client";

import Script from "next/script";
import type { CSSProperties } from "react";
import { useCallback, useState } from "react";

interface Finding {
  quote: string;
  issue: string;
  severity: "low" | "medium" | "high";
  suggestion: string;
  rulebookCitation: string;
}

type FindingStatus = "pending" | "applied" | "dismissed" | "not_found" | "error";

const SEVERITY_COLOR: Record<Finding["severity"], string> = {
  high: "#E4572E",
  medium: "#F2B134",
  low: "#4CAF50",
};

export default function TaskPane() {
  const [officeReady, setOfficeReady] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [findings, setFindings] = useState<Finding[]>([]);
  const [statuses, setStatuses] = useState<Record<number, FindingStatus>>({});
  const [expanded, setExpanded] = useState<Record<number, boolean>>({});

  const handleOfficeScriptLoad = useCallback(() => {
    Office.onReady(() => setOfficeReady(true));
  }, []);

  const runReview = useCallback(async () => {
    setLoading(true);
    setError(null);
    setFindings([]);
    setStatuses({});
    setExpanded({});

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
        const body = await response.json().catch(() => null);
        throw new Error(body?.error ?? `Review request failed (${response.status})`);
      }

      const data = (await response.json()) as { findings: Finding[] };
      setFindings(data.findings);
      setExpanded(data.findings.length > 0 ? { 0: true } : {});
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

  const dismissFinding = useCallback((index: number) => {
    setStatuses((prev) => ({ ...prev, [index]: "dismissed" }));
    setExpanded((prev) => ({ ...prev, [index]: false }));
  }, []);

  const toggleExpanded = useCallback((index: number) => {
    setExpanded((prev) => ({ ...prev, [index]: !prev[index] }));
  }, []);

  const total = findings.length;
  const resolvedCount = Object.values(statuses).filter(
    (status) => status === "applied" || status === "dismissed"
  ).length;

  return (
    <>
      <Script
        src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"
        onLoad={handleOfficeScriptLoad}
        strategy="afterInteractive"
      />
      <main style={styles.app}>
        <div style={styles.brandRow}>
          <span style={styles.brandMark}>§</span>
          <span style={styles.brandName}>Legal Copilot</span>
        </div>

        {total === 0 ? (
          <div style={styles.startPanel}>
            <button
              onClick={runReview}
              disabled={!officeReady || loading}
              style={{
                ...styles.primaryButton,
                ...((!officeReady || loading) && styles.buttonDisabled),
              }}
            >
              {loading ? "Reviewing..." : "Review Document Against Rulebook"}
            </button>
            {!officeReady && <p style={styles.mutedText}>Loading Office...</p>}
            {error && <p style={styles.errorText}>{error}</p>}
            {!loading && officeReady && !error && (
              <p style={styles.mutedText}>Run a review to see suggested edits here.</p>
            )}
          </div>
        ) : (
          <div style={styles.playbookPanel}>
            <div style={styles.playbookTitleRow}>
              <span style={styles.playbookTitle}>Rulebook Review</span>
              <button onClick={runReview} disabled={loading} style={styles.rerunButton}>
                {loading ? "Reviewing..." : "Re-run"}
              </button>
            </div>

            <div style={styles.progressRow}>
              <span style={styles.progressLabel}>{Math.min(resolvedCount + 1, total)}</span>
              <div style={styles.progressTrack}>
                <div
                  style={{
                    ...styles.progressFill,
                    width: `${(resolvedCount / total) * 100}%`,
                  }}
                />
              </div>
              <span style={styles.progressLabel}>{total}</span>
            </div>

            <ul style={styles.list}>
              {findings.map((finding, index) => {
                const status = statuses[index] ?? "pending";
                const isExpanded = !!expanded[index];
                const resolved = status === "applied" || status === "dismissed";

                return (
                  <li key={index} style={styles.card}>
                    <button onClick={() => toggleExpanded(index)} style={styles.cardHeader}>
                      <span
                        style={{
                          ...styles.dot,
                          background: resolved ? "#C7C7C7" : SEVERITY_COLOR[finding.severity],
                        }}
                      />
                      <span style={styles.cardTitle}>{finding.issue}</span>
                      <span style={styles.chevron}>{isExpanded ? "⌃" : "⌄"}</span>
                    </button>

                    {isExpanded && (
                      <div style={styles.cardBody}>
                        <p style={styles.citation}>{finding.rulebookCitation}</p>

                        <div style={styles.quoteBox}>
                          <blockquote style={styles.quoteText}>&ldquo;{finding.quote}&rdquo;</blockquote>
                        </div>

                        <div style={styles.suggestionBox}>{finding.suggestion}</div>

                        {status === "not_found" && (
                          <p style={styles.errorText}>
                            Could not locate this text in the document — try re-selecting it.
                          </p>
                        )}
                        {status === "error" && (
                          <p style={styles.errorText}>Something went wrong applying this comment.</p>
                        )}

                        <div style={styles.actionRow}>
                          <button
                            onClick={() => dismissFinding(index)}
                            disabled={resolved}
                            style={{ ...styles.dismissButton, ...(resolved && styles.buttonDisabled) }}
                          >
                            Dismiss
                          </button>
                          <button
                            onClick={() => applyFinding(finding, index)}
                            disabled={resolved}
                            style={{ ...styles.applyButton, ...(resolved && styles.buttonDisabled) }}
                          >
                            {status === "applied" ? "Applied" : "Apply"}
                          </button>
                        </div>
                      </div>
                    )}
                  </li>
                );
              })}
            </ul>

            <div style={styles.footer}>
              {resolvedCount} of {total} edits resolved
            </div>
          </div>
        )}
      </main>
    </>
  );
}

const styles: Record<string, CSSProperties> = {
  app: {
    fontFamily: "Segoe UI, sans-serif",
    background: "#fff",
    minHeight: "100vh",
    color: "#1a1a1a",
    fontSize: 13,
  },
  brandRow: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    padding: "14px 14px 10px 14px",
    borderBottom: "1px solid #EEE",
  },
  brandMark: {
    width: 22,
    height: 22,
    borderRadius: 6,
    background: "#1a1a1a",
    color: "#fff",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: 12,
    fontWeight: 700,
  },
  brandName: { fontSize: 14, fontWeight: 700 },
  startPanel: {
    padding: 16,
    display: "flex",
    flexDirection: "column",
    gap: 10,
  },
  primaryButton: {
    background: "#1a1a1a",
    color: "#fff",
    border: "none",
    borderRadius: 8,
    padding: "10px 16px",
    fontSize: 13,
    fontWeight: 600,
    cursor: "pointer",
  },
  buttonDisabled: { opacity: 0.5, cursor: "default" },
  mutedText: { fontSize: 12, color: "#888", margin: 0 },
  errorText: { fontSize: 12, color: "#C0392B", margin: "4px 0" },
  playbookPanel: { padding: "12px 0 16px 0" },
  playbookTitleRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "0 14px",
    marginBottom: 10,
  },
  playbookTitle: { fontSize: 15, fontWeight: 700 },
  rerunButton: {
    background: "transparent",
    border: "none",
    color: "#888",
    fontSize: 12,
    cursor: "pointer",
    textDecoration: "underline",
    padding: 0,
  },
  progressRow: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    padding: "0 14px",
    marginBottom: 14,
  },
  progressTrack: {
    flex: 1,
    height: 6,
    background: "#E8E8E8",
    borderRadius: 999,
    overflow: "hidden",
  },
  progressFill: {
    height: "100%",
    background: "linear-gradient(90deg, #F2994A, #27AE60)",
    borderRadius: 999,
    transition: "width 0.3s ease",
  },
  progressLabel: { fontSize: 11, color: "#888", minWidth: 12, textAlign: "center" },
  list: {
    listStyle: "none",
    margin: 0,
    padding: "0 14px",
    display: "flex",
    flexDirection: "column",
    gap: 8,
  },
  card: {
    border: "1px solid #E5E5E5",
    borderRadius: 10,
    overflow: "hidden",
  },
  cardHeader: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    width: "100%",
    padding: "12px 14px",
    background: "transparent",
    border: "none",
    textAlign: "left",
    fontSize: 13,
    fontWeight: 600,
    color: "#1a1a1a",
    cursor: "pointer",
  },
  dot: { width: 8, height: 8, borderRadius: "50%", flexShrink: 0 },
  cardTitle: { flex: 1 },
  chevron: { color: "#999", fontSize: 12 },
  cardBody: {
    padding: "0 14px 14px 14px",
    borderTop: "1px solid #F2F2F2",
    paddingTop: 10,
  },
  citation: { fontSize: 11.5, color: "#888", margin: "0 0 8px 0" },
  quoteBox: {
    background: "#F7F7F8",
    borderLeft: "3px solid #DDD",
    borderRadius: 6,
    padding: "8px 10px",
    marginBottom: 8,
  },
  quoteText: { margin: 0, fontStyle: "italic", fontSize: 12.5, color: "#555" },
  suggestionBox: {
    background: "#FFF4E5",
    border: "1px solid #F3D9B1",
    borderRadius: 8,
    padding: "10px 12px",
    fontSize: 12.5,
    color: "#4A3B22",
    lineHeight: 1.5,
    marginBottom: 10,
  },
  actionRow: { display: "flex", justifyContent: "flex-end", gap: 8 },
  dismissButton: {
    background: "#fff",
    border: "1px solid #D0D0D0",
    borderRadius: 6,
    padding: "6px 14px",
    fontSize: 12.5,
    cursor: "pointer",
    color: "#333",
  },
  applyButton: {
    background: "#1a1a1a",
    color: "#fff",
    border: "none",
    borderRadius: 6,
    padding: "6px 16px",
    fontSize: 12.5,
    fontWeight: 600,
    cursor: "pointer",
  },
  footer: {
    marginTop: 4,
    padding: "10px 14px 0 14px",
    fontSize: 12,
    color: "#888",
    textAlign: "center",
    borderTop: "1px solid #EEE",
  },
};
