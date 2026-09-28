import { readFile } from "fs/promises";
import path from "path";

const RULEBOOK_PATH = path.join(process.cwd(), "rulebook", "company-rulebook.md");

export async function loadRulebook(): Promise<string> {
  return readFile(RULEBOOK_PATH, "utf-8");
}
