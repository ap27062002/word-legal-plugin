import Anthropic from "@anthropic-ai/sdk";
import { loadRulebook } from "./rulebook";

export interface Finding {
  quote: string;
  issue: string;
  severity: "low" | "medium" | "high";
  suggestion: string;
  rulebookCitation: string;
}

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const REPORT_FINDINGS_TOOL: Anthropic.Tool = {
  name: "report_findings",
  description: "Report every place the document conflicts with, or should be revised to align with, the company rulebook.",
  input_schema: {
    type: "object",
    properties: {
      findings: {
        type: "array",
        items: {
          type: "object",
          properties: {
            quote: {
              type: "string",
              description: "The exact, verbatim text from the document this finding is about. Must match the source text exactly so it can be located.",
            },
            issue: { type: "string", description: "What conflicts with policy, in one or two sentences." },
            severity: { type: "string", enum: ["low", "medium", "high"] },
            suggestion: { type: "string", description: "Concrete replacement or added language that would resolve the issue." },
            rulebookCitation: { type: "string", description: "The rulebook section or clause this finding is based on." },
          },
          required: ["quote", "issue", "severity", "suggestion", "rulebookCitation"],
        },
      },
    },
    required: ["findings"],
  },
};

export async function reviewDocument(documentText: string): Promise<Finding[]> {
  const rulebook = await loadRulebook();

  const message = await client.messages.create({
    model: "claude-sonnet-5",
    max_tokens: 4096,
    system:
      "You are a legal compliance reviewer for this company. You are given the company's rulebook and the full text of a document currently open for editing. " +
      "Identify every clause or passage in the document that conflicts with, or should be revised to align with, the rulebook. " +
      "Be precise: only flag real conflicts or gaps, not stylistic preferences. Quote the document text exactly as written so it can be located. " +
      "Report your findings using the report_findings tool.",
    tools: [REPORT_FINDINGS_TOOL],
    tool_choice: { type: "tool", name: "report_findings" },
    messages: [
      {
        role: "user",
        content: `<company_rulebook>\n${rulebook}\n</company_rulebook>\n\n<document>\n${documentText}\n</document>`,
      },
    ],
  });

  const toolUse = message.content.find((block) => block.type === "tool_use");
  if (!toolUse || toolUse.type !== "tool_use") {
    return [];
  }

  const input = toolUse.input as { findings: Finding[] };
  return input.findings ?? [];
}
