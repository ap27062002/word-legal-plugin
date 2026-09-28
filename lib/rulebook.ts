import playbook from "../rulebook/chargepoint-nda-playbook.json";

interface PlaybookRule {
  id: string;
  title: string;
  startingPosition: string;
  notAcceptable: string[];
  fallbacks: string[];
}

export function loadRulebook(): string {
  return playbook.rules
    .map((rule: PlaybookRule) => {
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
