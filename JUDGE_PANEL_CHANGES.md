# Judging panel update — what changed & where

Changes to how the panel arbitrates disagreements:

1. **NotebookLM match-authority (existence vs. validity)** — NotebookLM is treated
   as authoritative on *whether an existing attribute match exists and which one it
   is* (it's almost never wrong there; it beats Gemini outright on that). Whether
   the attribute should actually be *merged* into that confirmed target is still a
   genuine debate with Gemini under the Akeneo rules — a real match can still be a
   bad merge, so Merge is not auto-accepted just because NotebookLM proposed it.
2. **Judgment info in its own columns; Notes stays yours** — arbitration never
   writes to the `Notes` column (that is exclusively for your own action notes).
   It writes the outcome to a dedicated `Panel Verdict` column, always records
   each judge's *initial* suggestion (`NLM Verdict`, `Gemini Verdict`), and stores
   the round-by-round transcript in `Debate Log`, shown as a collapsible "How the
   debate went" section on the card (expand to review, collapse when done).
3. **Multi-round debate (hybrid)** — instead of one arbitration call, Claude now
   debates each disagreement over up to 3 rounds. Anything still split after
   round 3 gets a ready-to-paste "fresh opinion" prompt for a new NotebookLM /
   Gemini chat (shown on the Judges tab).
4. **Dictionary enrichment** — each debated case now carries the *actual*
   dictionary record(s) of the target attribute(s) (datatype, dropdown/option
   list, unit, sample values, description) so Claude can apply the merge-validity
   rules concretely. Only the referenced targets are injected, not the whole
   dictionary — existence stays NotebookLM's call, validity gets real evidence.
5. **Trail attached to already-decided rows (never overwrites your decision)** —
   arbitration used to *skip* any row that already had a `Your Decision`, so those
   rows never got judging notes. Now it still leaves your decision untouched but
   attaches the panel trail (`NLM Verdict`, `Gemini Verdict`) and a comparison
   `Panel Verdict`: `✅ Panel unanimous (…) — matches your decision`,
   `⚠️ Panel unanimous: … — you chose …` when the panel would have gone another
   way, or `⚖ Judges split — NLM: … | Gemini: … — you chose …`. This is the
   **cheap path** — no Claude debate is run for rows you've already settled; you
   get the two judges' positions and a flag. To force a full debate on a settled
   row, clear its `Your Decision` cell and re-run.

---

## Frontend (`index.html`) — DONE (already edited in this repo)

- `isJudged()` now keys off the dedicated `Panel Verdict` column (not `Notes`).
- Review cards render `Panel Verdict`, `NLM Verdict` (initial), `Gemini Verdict`
  (initial), and a collapsible `Debate Log` ("How the debate went (N rounds)").
- Judges page has a "Still split after 3 rounds — get a fresh opinion" section,
  loaded by the new `loadDebateReprompt()` (calls `action=getDebateReprompt`).
- Arbitration button copy updated to mention the 3-round debate.
- The `Notes` field on each card is untouched by arbitration — yours to fill.
- The arbitration trail (`Panel Verdict` + initial suggestions + `Debate Log`)
  shows on **all three filters** (Pending / All / Decided) and does **not**
  disappear when a row becomes decided.
- **Trail is no longer tied to save status.** The `Debate Log` renders **fully
  expanded by default** regardless of whether the row is undecided, human-saved,
  or auto-filled by arbitration. It folds **only when you tap it**, and that
  manual fold is remembered in a `trailCollapsed` set (`onTrailToggle`) so
  re-renders — save, autosave, filter switch, the 4s log poll — can't re-open or
  re-collapse it. A page reload resets everything to expanded for a fresh audit.
  (`humanSaved` now only drives the "✓ saved" badge, never folding.)

## Backend (Apps Script) — apply these

Easiest path: open the Apps Script editor, select all, and paste the full
contents of **`copd_code_updated.gs`** (in this repo). That file already has
every change below baked in and verified.

If you'd rather patch by hand, here is exactly what changed:

| # | Where | Change |
|---|-------|--------|
| A | `SHEETS` object | Added `DEBATE_REPROMPT: 'DEBATE_REPROMPT'` |
| B | after `getSheet()` | Added `getOrCreateSheet()` helper |
| C | `doGet()`, before the "Unknown action" return | Added `action === 'getDebateReprompt'` handler |
| D | `runJudgeArbitration()` column setup | `ensureReviewColumns(['Panel Verdict','NLM Verdict','Gemini Verdict','Debate Log'])` + `cVerdict`/`cNLM`/`cGem`/`cDebate` |
| E | `runJudgeArbitration()` missing-judge branch | Writes each judge's verdict even when one is missing |
| F | `runJudgeArbitration()` arbitration block | Replaced single call with `debateDisagreements()` loop, column-expanding write-back, and hybrid re-prompt write. **Writes outcome to `Panel Verdict`, never `Notes`** (Notes is reviewer-only); unanimous rows also store both judges' initial suggestions |
| G | end of file | Old `arbitrateDisagreements` renamed `_DEPRECATED` (unused); added the **MULTI-ROUND DEBATE MODULE**: `DEBATE_MAX_ROUNDS`, `cap`, `fmtJudge`, `ensureReviewColumns`, `debateDisagreements`, `runDebateRound`, `writeDebateReprompts`, `clearDebateReprompts` |
| H | debate module + `runJudgeArbitration` | Dictionary enrichment: `buildDictIndex`, `lookupDictRecords`, `fmtDictRecord`; each disagreement carries `targetRecords`, injected into the debate prompt |
| I | `runJudgeArbitration()` main loop | Replaced the early "already decided → skip" guard with an `alreadyDecided` flag. Both-judge rows now always write `NLM Verdict`/`Gemini Verdict`; already-decided rows get a compare-and-flag `Panel Verdict` and are **not** re-debated (decision left untouched). Uses the existing top-level `cap()` helper |

### Manual steps outside the code

- **No sheet setup needed.** The `DEBATE_REPROMPT` tab is auto-created, and the
  four new REVIEW columns (`Panel Verdict`, `NLM Verdict`, `Gemini Verdict`,
  `Debate Log`) are auto-added by `ensureReviewColumns` on the next arbitration run.
- **Heads-up on old rows:** the `Notes` column is no longer written by
  arbitration, but any arbitration text left in `Notes` from *previous* runs stays
  until you clear it. New runs write to `Panel Verdict` instead.
- Redeploy the Apps Script Web App so the new `getDebateReprompt` action is live
  (Deploy → Manage deployments → edit → new version), since the frontend calls it.

### Tuning knobs

- `DEBATE_MAX_ROUNDS` (default `3`) — number of debate rounds before a case is
  handed to you with a fresh-opinion prompt.
- Match-authority wording lives in the `JUDGE AUTHORITY` paragraph inside
  `runDebateRound`'s prompt. It currently makes NotebookLM authoritative on match
  *existence* while leaving merge *validity* open to debate; adjust there if needed.
