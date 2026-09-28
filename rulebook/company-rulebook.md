# Company Rulebook (placeholder)

Replace this file with your actual company policy / rulebook content.

For the MVP, the entire contents of this file are sent to the LLM as
context on every review request (see `lib/agent.ts`), so keep it under
your model's context window. If the real rulebook is large or changes
often, swap `lib/rulebook.ts` for a retrieval step (embed + chunk the
rulebook, fetch only the relevant sections per review) instead of full
context stuffing.

## Example section: Confidentiality

All contracts must include a confidentiality clause covering both
parties, with a minimum term of 2 years post-termination.

## Example section: Governing Law

All agreements must specify [Your State/Country] as the governing law
and venue for disputes.

## Example section: Limitation of Liability

Liability caps below 12 months of fees paid require Legal sign-off
before the document can be sent externally.
