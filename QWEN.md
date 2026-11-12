# Repository agent pipeline — contract

This repository is developed by an autonomous multi-agent pipeline.
Every session (web, TUI, or headless) must follow this contract. The
stages run as subagents defined in the shared user-level
`~/.qwen/agents/` (a repo-local `.qwen/agents/` overrides if present);
the session acting as Engineering Lead dispatches them per
`~/.qwen/lead.md` (same override rule). All stage artifacts
(`.agents/`, `docs/`) are always repo-local.

The stage table below is the sole source of stage names, numbers,
artifacts and model assignments — skills and agent files reference it,
they do not restate it, so they cannot drift from it.

## Pipeline (in order)

| # | Stage | Artifact | Model |
|---|-------|----------|-------|
| 1 | Research | `.agents/research.md` | qwen-dense |
| 2 | Architect | `.agents/architecture.md` | qwen-dense |
| 3 | Reviewer | `.agents/review.md` | qwen-dense |
| 4 | Product Manager | `.agents/requirements.md` (exhaustive work list) | qwen-dense-small |
| 5 | Coding | `.agents/implementation.md` | qwen-dense |
| 6 | QA | `.agents/qa.md` | qwen-dense |
| 7 | Security | `.agents/security.md` | qwen-dense |
| 8 | DevOps | `.agents/ops.md` | qwen-dense |
| 9 | Documentation | `docs/` | qwen-dense-small |
| 10 | Hindsight | `.agents/hindsight.md` | qwen-dense-small |

Terminal (read-only, no artifact): Reporter delivers a one-screen
status report in chat only. Stage 6 and 7 may run in parallel after
Coding completes.

## Artifact rules

- The `.agents/` files are the source of truth. Chat carries summaries
  only; never rely on a chat message for state a file should record.
- Each stage reads its predecessor artifacts and writes exactly its own
  artifact. Do not ask an agent for anything it already wrote to a file.
- Update `.agents/implementation.md` after every work-list item
  (status, files touched, validation) — not in a batch.
- On context loss or summarisation, never guess what is already done:
  re-read `.agents/pipeline.md` and the stage artifacts, then resume at
  the first uncompleted item. A partial pass is a failed pass.
- Keep artifacts lean: at most ~200 lines each. Artifacts are re-read
  by every downstream stage, so a fat file taxes the whole run; when a
  file approaches the cap, the owning stage compresses settled material
  into a short summary section before writing the rest.

## Findings

Every QA/Security finding is classified:

- **Blocking**: breaks specified behaviour, or exploitable / critical
  design flaw. Must be fixed before the stage passes.
- **Risk-accepted**: documented and carried forward; never blocks the
  pipeline; flows forward as documented debt.

Design-level security findings (trust boundaries, authentication, data
flow) are marked explicitly in `.agents/security.md` — they route the
rework loop back to the Architect.

## Rework loops

- QA Blocking: Coding fix pass → QA post-fix verification → continue.
- Security Blocking, code-level: Coding fix pass → Security post-fix
  verification → continue.
- Security Blocking, design-level: Architect → continue from Reviewer.
- Cap each loop at two fix-verify rounds. On the third round, STOP:
  write a "HUMAN DECISION REQUIRED" section with concrete options into
  `.agents/pipeline.md`, leave the goal incomplete, and report the
  options in chat. Never keep looping.

## State ledger

`.agents/pipeline.md` is kept current by the lead only:

- The objective and what "done" means
- Per-stage status: done / running / blocked / escalated
- Rework round counts per gate
- Open Blocking vs risk-accepted findings

Before acting, the lead determines: which specialists have completed,
which artifacts are missing, whether `qa.md` / `security.md` contain
unresolved Blocking findings, and which specialist runs next — then
updates the ledger. Explain why a stage is skipped or repeated.

## Shared engineering rules (all stages)

- Follow the repository's existing languages, conventions and patterns;
  keep diffs minimal. No design changes in a fix pass.
- Do not leave TODOs.
- The repository maintains a GLOSSARY.md of domain vocabulary: Research
  lists candidate terms, the Architect owns and finalises the file
  before the Coding pass, and all stages use those terms in names
  and prose.
- Non-trivial design decisions are recorded as one-page ADRs in
  docs/adr/ (context, decision, alternatives considered, consequences),
  written by the stage that makes the decision; Hindsight ingests them.
- Verify with the project's own checks (tests, lint, build), not by
  eyeballing.
- Where a configured MCP server covers the relevant platform (Azure,
  Bicep/IaC, current official docs), verify with it, not model memory —
  a plan checked against live tooling beats one recalled from training.
- One edit per file per step; on edit failure, re-read the file and
  retry.
- If a required package is unavailable (offline, no feed, version
  missing): do not retry in a loop. Record the exact command a human
  should run, then continue with the remaining items.
- Read-only stages (Research, Reviewer, Reporter) never modify the
  working tree; the lead must keep them in plan/approvalMode:plan.
- Deliverables must include the validation and rollback story where
  relevant (Docs comes from DevOps's `.agents/ops.md`).

## Done means

All ten artifacts exist, no open Blocking findings (every one resolved
or risk-accepted), the Reporter's one-screen status report has been
delivered in chat, and `.agents/pipeline.md` reflects this.