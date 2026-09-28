import { readFile } from "fs/promises";
import path from "path";

const PLAYBOOK_PATH = path.join(process.cwd(), "rulebook", "chargepoint-nda-playbook.json");

interface PlaybookRule {
  id: string;
  title: string;
  startingPosition: string;
  notAcceptable: string[];
  fallbacks: string[];
}

interface Playbook {
  rules: PlaybookRule[];
}

export async function loadRulebook(): Promise<string> {
  const raw = await readFile(PLAYBOOK_PATH, "utf-8");
  const playbook = JSON.parse(raw) as Playbook;

  return playbook.rules
    .map((rule) => {
      const notAcceptable = rule.notAcceptable.map((item) => `  - ${item}`).join("\n");
      const fallbacks = rule.fallbacks.map((item) => `  - ${item}`).join("\n");

      return [
        `## ${rule.title} (rule id: ${rule.id})`,
        `Starting position: ${rule.startingPosition}`,
        `Not acceptable:`,
        notAcceptable,
        `Fallbacks (acceptable if the starting position can't be held):`,
        fallbacks,
      ].join("\n");
    })
    .join("\n\n");
}
