# python-pptx-extended — Contributor & Agent Guide

This is `MHoroszowski/python-pptx`, a fork of `scanny/python-pptx` published on PyPI as `python-pptx-extended`. Commit `253dbc87` on master (PR #30 — transparency & opacity, issue #17) is the canonical example of a feature PR in this repo.

This file documents project rules that apply to every contributor — human or AI — working in this codebase. Claude Code auto-merges it into any session opened in this directory.

## 1. Editable install resolves to the path it was installed against

Run `pip show python-pptx-extended` (or `python-pptx`, depending on what's installed) to see the package location. With an editable install (`pip install -e .`), that location points at the clone the install was run against — typically THIS directory's `src/`.

**Consequence for `git worktree`:** pytest in a parallel worktree still imports from the original install location, not the worktree's `src/`. Edits in the worktree are invisible to the test runner. Tests appear to "fail mysteriously" against the new code because the new code never loads.

**Two acceptable workflows:**

- **Single feature at a time:** create a feature branch directly in the main checkout. No worktree. Editable install Just Works.
- **Parallel checkouts via worktree:** every pytest invocation MUST set `PYTHONPATH=src`, e.g. `PYTHONPATH=src python3 -m pytest tests/ -q`. No exceptions.

If you skip this and run plain `pytest` from a worktree, you will spend hours debugging "test failures" that are actually phantom imports.

## 2. Manual port over `git cherry-pick` for upstream fixes

This fork's master had a repo-wide `ruff format` pass (PR #10). Upstream `scanny/python-pptx` master did not. `git cherry-pick` of upstream PRs will conflict on whitespace across nearly every touched file. **Always manually apply the semantic changes from upstream PRs onto the ruff-formatted base.** Reuse the design and tests from upstream; redo the diff against current master.

## 3. Test discipline (non-negotiable for any feature PR)

- **Unit tests:** new tests in `tests/` mirroring source layout. Aim for ≥30 new tests for a meaningful surface (PR #30 added 31 in `tests/dml/test_transparency.py`).
- **Behave acceptance:** `python3 -m behave features/ --no-color 2>&1 | tail -5` must show `0 failed` (current baseline: 981 scenarios).
- **Pytest:** `python3 -m pytest tests/ -q` must show `0 failed` (current baseline: 3017 passed on master, ef8ee6a8).
- **Ruff:** both `ruff format src tests` (no diff or applied cleanly) and `ruff check src tests` (All checks passed).

Run pytest at least once during implementation, not just at the end. TDD ordering is the most reliable way to get there: write the first failing test, run it, see it fail with the expected error, implement, see it pass, then expand. Writing everything before running anything is how subtle indentation bugs and stale-import gotchas turn into hour-long debugging sessions.

## 4. Tooling

- **Ruff:** install via `pip install ruff` if it isn't on PATH already. Locate with `which ruff`. Used for both lint (`ruff check`) and format (`ruff format`).
- **Pytest, behave:** invoke as `python3 -m pytest tests/` and `python3 -m behave features/`.
- **`gh` CLI:** target this fork as `MHoroszowski/python-pptx`.

## 5. Commit message format

Reference `git log 253dbc87 --format=%B -n1` for the canonical shape: scope, design notes, test counts, `Closes #N` (or `Refs #N` if partial). Long-form, not terse.

## 6. Branch / PR workflow

1. `git checkout -b feature/<slug>` off master
2. Implement → test → ruff → commit
3. Drop a `uat_<slug>.py` at the repo root (untracked) — small Python script that builds a `.pptx` exercising the new API and prints a round-trip read-back. PR #30 used this for issue #17 review.
4. The maintainer runs the UAT, opens the `.pptx` in PowerPoint or Keynote, and gives signoff.
5. Then push the branch and open a PR.

**Automated runs must not push or open PRs without maintainer approval.** The approval-gated UAT step is the rule, not a suggestion.

## 7. AI agent reporting contract

When an AI agent or automated workflow finishes implementation work in this repo, the wrap-up report MUST include the literal output of:

```
python3 -m pytest tests/ -q | tail -3
ruff check src tests | tail -3
python3 -m behave features/ --no-color 2>&1 | tail -3
```

Pasted verbatim. If any shows a failure, the agent stops there and reports `verification-failed` rather than committing. Self-attestation ("tests pass") without the captured output is not acceptable — too easy to skip the actual run and inherit a false sense of done.
