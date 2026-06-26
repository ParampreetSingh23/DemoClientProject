# AGENTS.md

Project-level instructions for AI coding agents working in this repository.
These rules apply on top of any global harness guidance and any user
instructions given in the active session.

---

## 1. Project context

- **Purpose:** browser-based reconciliation app that matches payment gateway
  aggregator Excel exports (HDFC, BillDesk, …) against API invoice exports
  (CME, CMS, LMA, YLC) and produces a Tally-ready reconciled workbook.
  See `RECONCILIATION_DOCUMENTATION.md` for the full functional spec.
- **Stack:** Vite + React 19, single-page app. Source lives in `src/`
  (`main.jsx`, `styles.css`). Excel I/O via the `xlsx` package.
- **Build / run:** `npm run dev`, `npm run build`, `npm run preview`.
- **Tests:** Playwright is installed as a dev dependency; tests are not
  checked in yet — confirm with the user before adding or running them.
- **Ignored paths:** `node_modules/`, `dist/`, `outputs/`, all `*.xlsx` /
  `*.xls` / `*.csv` / `*.ndjson` files, `*.log`, and `.env*`. Do not commit
  any of these.
- **Repo state:** `master` branch, single initial commit. Detect the default
  branch dynamically rather than hardcoding `main` / `master`.

---

## 2. Git

- Stage files explicitly by name (e.g. `git add path/to/file`). Avoid
  `git add -A` and `git add .` — they sweep up untracked secrets, build
  artefacts, and stray files outside the change.
- Never run destructive commands (`git reset --hard`, `git push --force`,
  `git branch -D`, `git clean -f`) on `main`, `master`, `release/*`, or
  other protected branches without explicit user approval.
- Prefer creating new commits over amending published commits. Only amend
  when the user explicitly asks.
- Never skip hooks (`--no-verify`) or bypass signing unless the user
  explicitly asks. If a hook fails, fix the underlying issue.
- When running automated git commands that may invoke an editor (e.g.
  `git rebase`, `git commit`, `git merge --squash`), set
  `GIT_EDITOR=true` — an interactive shell must not block execution or
  cause the command to hang.
- Do not hardcode branch names like `main` or `master`. Detect the default
  branch dynamically, e.g.:

  ```bash
  git symbolic-ref refs/remotes/origin/HEAD --short | sed 's/origin\///'
  ```

  Use the detected name in scripts and commands.

---

## 3. `gh` CLI

`gh` is the canonical interface for GitHub. Prefer it over scraping web
URLs or guessing API paths. Discover flags with `gh <cmd> --help` rather
than enumerating here.

### 3.1 Auth

- `gh auth status` — confirm login before anything that talks to GitHub.
- If logged out, ask the user to run `gh auth login`.

### 3.2 Repo targeting

- Repo is inferred from cwd. Pass `-R OWNER/REPO` when running outside
  the repo.

### 3.3 PR review — non-obvious bits

- Find PRs awaiting your review:

  ```bash
  gh pr list --search "review-requested:@me"
  ```

- Existing review state — two endpoints, easy to confuse:

  ```bash
  gh api repos/OWNER/REPO/pulls/123/comments  --paginate   # inline, line-anchored
  gh api repos/OWNER/REPO/issues/123/comments --paginate   # PR-level conversation
  ```

- Post inline comments in one review (line-anchored, multi-comment) — no
  `gh pr review` flag for this; use the API:

  ```bash
  gh api repos/OWNER/REPO/pulls/123/reviews -f event=COMMENT \
    -f body="overall notes" \
    -F 'comments[][path]=src/foo.py' -F 'comments[][line]=42' \
    -F 'comments[][body]=this is wrong because…'
  ```

- Resolve a review thread (GraphQL):

  ```bash
  gh api graphql -f query='mutation($id:ID!){resolveReviewThread(input:{threadId:$id}){thread{isResolved}}}' \
    -F id=THREAD_NODE_ID
  ```

- Reply to a specific inline thread:

  ```bash
  gh api repos/OWNER/REPO/pulls/123/comments/COMMENT_ID/replies -f body="fixed in abc1234"
  ```

- Top-level review verbs: `gh pr review <N> --approve|--request-changes|--comment -b "…"`
  (see §3.6 before posting).

### 3.4 Workflow runs

- Default to **failed-only** logs, never the full log:

  ```bash
  gh run view 123456 --log-failed          # preferred
  gh run view 123456 --log | tail -200     # only if --log-failed isn't enough
  ```

- Find the run behind a PR's latest push:

  ```bash
  gh pr checks 123 --json name,state,link,workflow
  ```

### 3.5 `gh api` cheatsheet

- `-f key=val` — string param.
- `-F key=val` — typed (numbers, booleans, `@file`).
- `-X METHOD` — HTTP verb.
- `--jq '.field'` — filter response.
- `--paginate` — follow `Link` headers.

### 3.6 Never without explicit consent

Anything that publishes, mutates, or notifies needs an explicit
in-conversation request from the user. Do **not** run these unprompted:

- Posting to a PR/issue: `gh pr review` (any of `--approve`,
  `--request-changes`, `--comment`), `gh pr comment`, `gh issue comment`,
  posts via `gh api .../comments` or `.../reviews`.
- State changes on PRs: `gh pr merge` (any flags), `gh pr close`,
  `gh pr reopen`, `gh pr ready` (and `--undo`), `gh pr edit`.
- CI: `gh run rerun`, `gh run cancel`.
- Issues: `gh issue close`, `gh issue reopen`, `gh issue edit`,
  `gh issue delete`.
- Releases: `gh release create`, `gh release edit`, `gh release delete`.
- Any `gh api -X POST/PATCH/PUT/DELETE` that mutates state, including
  resolving review threads.
- Git remote ops: pushing branches, force-push, deleting branches/tags.

Read-only commands (`list`, `view`, `diff`, `checks`, `status`,
`gh api` GETs) are fine. When in doubt, surface the command in chat and
wait for confirmation.

### 3.7 Output discipline

- `gh run view --log` is huge — prefer `--log-failed` or pipe through
  `| tail -N`.
- `gh api ... --paginate` can be massive — add `--jq` to narrow it down.
- `gh pr diff` on big PRs — use `--name-only` first, then targeted reads.
