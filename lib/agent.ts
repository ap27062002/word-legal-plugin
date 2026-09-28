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
  description: "Report every place the document deviates from the negotiation playbook, ranked by how far it falls from an acceptable position.",
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
              description:
                "The exact, verbatim text from the document this finding is anchored to, so it can be located and commented on. " +
                "If the issue is a missing clause (the playbook rule's topic doesn't appear in the document at all), quote the nearest existing heading or sentence the new clause should follow.",
            },
            issue: {
              type: "string",
              description:
                "What's wrong, in one or two sentences: how the document's language compares to the rule's starting position, and whether it lands in the rule's notAcceptable list or is an unlisted deviation.",
            },
            severity: {
              type: "string",
              enum: ["low", "medium", "high"],
              description:
                "high = matches or is equivalent to something in the rule's notAcceptable list, or a required clause is missing entirely. " +
                "medium = deviates from the starting position and isn't covered by any listed fallback (needs a judgment call). " +
                "low = deviates from the starting position but matches one of the rule's listed fallbacks (acceptable, flagged for visibility only).",
            },
            suggestion: {
              type: "string",
              description:
                "Concrete replacement or added language, drawn from the rule's starting position (or nearest acceptable fallback) that would resolve the issue.",
            },
            rulebookCitation: {
              type: "string",
              description: "The playbook rule this finding is based on, as \"<title> (rule id: <id>)\".",
            },
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
      "You are a legal negotiator reviewing a document against this company's negotiation playbook. Each playbook rule has a startingPosition " +
      "(what to open with), a notAcceptable list (positions that must never be accepted as-is), and a fallbacks list (positions that are fine to " +
      "settle on if the starting position can't be held). For every rule, check what the document actually says on that topic (including if it's " +
      "missing entirely) and compare it against these three. Only flag real deviations, not stylistic differences that don't change the legal " +
      "position. Quote the document text exactly as written so it can be located. Report your findings using the report_findings tool.",
    tools: [REPORT_FINDINGS_TOOL],
    tool_choice: { type: "tool", name: "report_findings" },
    messages: [
      {
        role: "user",
        content: `<negotiation_playbook>\n${rulebook}\n</negotiation_playbook>\n\n<document>\n${documentText}\n</document>`,
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
