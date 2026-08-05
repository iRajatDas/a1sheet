<!-- dgc-policy-v11 -->
# Dual-Graph Context Policy

This project uses a local dual-graph MCP server for efficient context retrieval.

## MANDATORY: Always follow this order

1. **Call `graph_continue` first** — before any file exploration, grep, or code reading.

2. **If `graph_continue` returns `needs_project=true`**: call `graph_scan` with the
   current project directory (`pwd`). Do NOT ask the user.

3. **If `graph_continue` returns `skip=true`**: project has fewer than 5 files.
   Do NOT do broad or recursive exploration. Read only specific files if their names
   are mentioned, or ask the user what to work on.

4. **Read `recommended_files`** using `graph_read` — **one call per file**.
   - `graph_read` accepts a single `file` parameter (string). Call it separately for each
     recommended file. Do NOT pass an array or batch multiple files into one call.
   - `recommended_files` may contain `file::symbol` entries (e.g. `src/auth.ts::handleLogin`).
     Pass them verbatim to `graph_read(file: "src/auth.ts::handleLogin")` — it reads only
     that symbol's lines, not the full file.
   - Example: if `recommended_files` is `["src/auth.ts::handleLogin", "src/db.ts"]`,
     call `graph_read(file: "src/auth.ts::handleLogin")` and `graph_read(file: "src/db.ts")`
     as two separate calls (they can be parallel).

5. **Check `confidence` and obey the caps strictly:**
   - `confidence=high` -> Stop. Do NOT grep or explore further.
   - `confidence=medium` -> If recommended files are insufficient, call `fallback_rg`
     at most `max_supplementary_greps` time(s) with specific terms, then `graph_read`
     at most `max_supplementary_files` additional file(s). Then stop.
   - `confidence=low` -> Call `fallback_rg` at most `max_supplementary_greps` time(s),
     then `graph_read` at most `max_supplementary_files` file(s). Then stop.

## Token Usage

A `token-counter` MCP is available for tracking live token usage.

- To check how many tokens a large file or text will cost **before** reading it:
  `count_tokens({text: "<content>"})`
- To log actual usage after a task completes (if the user asks):
  `log_usage({input_tokens: <est>, output_tokens: <est>, description: "<task>"})`
- To show the user their running session cost:
  `get_session_stats()`

Live dashboard URL is printed at startup next to "Token usage".

## Rules

- Do NOT use `rg`, `grep`, or bash file exploration before calling `graph_continue`.
- Do NOT do broad/recursive exploration at any confidence level.
- `max_supplementary_greps` and `max_supplementary_files` are hard caps - never exceed them.
- Do NOT dump full chat history.
- Do NOT call `graph_retrieve` more than once per turn.
- After edits, call `graph_register_edit` with the changed files. Use `file::symbol` notation (e.g. `src/auth.ts::handleLogin`) when the edit targets a specific function, class, or hook.

## Context Store

Whenever you make a decision, identify a task, note a next step, fact, or blocker during a conversation, call `graph_add_memory`.

**To add an entry:**
```
graph_add_memory(type="decision|task|next|fact|blocker", content="one sentence max 15 words", tags=["topic"], files=["relevant/file.ts"])
```

**Do NOT write context-store.json directly** — always use `graph_add_memory`. It applies pruning and keeps the store healthy.

**Rules:**
- Only log things worth remembering across sessions (not every minor detail)
- `content` must be under 15 words
- `files` lists the files this decision/task relates to (can be empty)
- Log immediately when the item arises — not at session end

## Session End

When the user signals they are done (e.g. "bye", "done", "wrap up", "end session"), proactively update `CONTEXT.md` in the project root with:
- **Current Task**: one sentence on what was being worked on
- **Key Decisions**: bullet list, max 3 items
- **Next Steps**: bullet list, max 3 items

Keep `CONTEXT.md` under 20 lines total. Do NOT summarize the full conversation — only what's needed to resume next session.

---

# a1sheet — engineering rules

Binding for all work in this repo. Not aspirations: a change that violates a MUST
is wrong even if it passes tests.

## 0. The rule that overrides the rest

**Compose, do not configure.** shadcn/ui is the target pattern — the consumer
assembles primitives and owns the layout; the library never does.

```tsx
// ✅ composition — the consumer owns the tree
<Sheet.Root workbook={wb}>
  <Sheet.Toolbar />
  <Sheet.Grid />
  <MyFooter />          {/* their component, our data, via useSheet() */}
</Sheet.Root>

// ❌ configuration — every new arrangement needs a new prop
<Spreadsheet showToolbar showStatusBar footer="custom" />
```

**Never add a `show*` / `enable*` / `hide*` boolean prop.** If something is
optional it is a child the consumer chooses not to render. A boolean that gates a
subtree is a composition failure, and boolean parameters are separately banned (§3).

Sanctioned exception: a **preset** that composes primitives in a documented default
arrangement and takes no layout props at all. Presets are convenience, never the
extension point.

## 1. Architecture

- **Headless first.** Behavior lives in hooks and framework-agnostic modules; a
  component is a thin renderer over a hook. Logic that cannot run without
  rendering is in the wrong place.
- **Two entrypoints, permanently.** `"."` is framework-agnostic and MUST NOT reach
  React. `"./react"` is the only place React appears. Enforced by a test.
- **Zero runtime dependencies.** `dependencies` stays `{}`. A feature that seems to
  need a package gets a minimal hand-written implementation. Enforced by a test.
- **Tree-shakeable, side-effect free.** `sideEffects: false` holds. No module-level
  work beyond `const` tables. No singletons, no global mutable state.
- **One responsibility per module.** If describing a file needs "and", split it.

## 2. React

MUST:
- **Context, not prop drilling.** Primitives read from context, never an `api`
  prop. One provider; a primitive used outside it throws a named error.
- **`asChild` on every leaf primitive** rendering a single DOM element.
- **`forwardRef` + `useImperativeHandle`** on anything a consumer may focus,
  scroll, or measure.
- **Stable callbacks and keys.** `useCallback` across component boundaries; never
  an array index as a key.
- **Controlled and uncontrolled both.** `value` + `onValueChange`, or
  `defaultValue`. Never silently switch modes.
- **`"use client"`** on every module under `./react`. RSC-safe by default.
- **SSR-safe.** No `window`/`document` at module scope or during render.

MUST NOT:
- Render props where a slot or `asChild` works.
- `useEffect` to derive state computable during render.
- A component that both mutates and renders.

## 3. Code quality

Banned: magic numbers and strings, hardcoded values, dead or commented-out code,
unused exports, circular dependencies, global mutable state, god objects, deep
nesting (early-return instead), speculative generality.

- **No boolean parameters.** `fn(true)` is unreadable — named options object or
  distinct functions.
- **>3 parameters** becomes an options object.
- **Consistent naming.** Same concept, same word. `row`/`col` are always
  zero-indexed; anything 1-indexed says so in its name.
- **YAGNI beats extensibility.** No plugin seam until a second real consumer
  exists. Premature abstraction costs as much as none.
- **Immutability.** Model functions take a value and return a new one. Only
  `useWorkbook` clones-on-write.

## 4. TypeScript

- `strict` + `noUncheckedIndexedAccess`. No `any` — use `unknown` and narrow.
- Discriminated unions over optional-field soup; `never` for exhaustiveness.
- `readonly` on anything consumers should not mutate; `as const` for literal tables.
- Branded types where values must not be confused (`CellKey`).
- **Typed errors, not bare strings.** Every throw is a named class with a stable
  `code`. Consumers branch on the code, never on message text.
- Public types are API: exported, documented, versioned.

## 5. Errors and failure

- **Fail fast on programmer error** (bad arguments, primitive outside its
  provider) — throw immediately, message names the fix.
- **Be forgiving with user data.** A bad formula degrades to an error value in that
  one cell and never breaks the sheet. A bad file throws one clear, actionable
  error rather than half-importing.
- Messages state what was wrong, what was expected, what to do next.

## 6. Performance

- **Virtualize anything unbounded.**
- **Long work is async, cancellable, and reports progress.** Parsing above a size
  threshold takes an `AbortSignal` and `onProgress`. Chunk it or hand it to a
  Worker; never block the main thread on a 38 MB file.
- Memoize deliberately, with a stated reason. Unexplained `useMemo` is noise.
- Measure before optimizing — and do not leave obvious waste either.

## 7. Security

- **CSV formula injection is mitigated on export.** Values starting with
  `= + - @`, tab, or CR are neutralized. Real RCE vector, not theoretical.
- Validate at the boundary, sanitize on output, never build HTML from untrusted
  strings.
- Sanitize download filenames: no separators, no traversal.
- Secure by default; unsafe behavior is opt-in and documented.

## 8. Testing

- Every bug fix ships a regression test that fails without the fix.
- Test through the public API, not internals.
- **A render test is not a behavior test.** Component work needs interaction tests
  that type, click, and navigate.
- Real-world fixtures for file I/O — self-round-trip proves only self-consistency.
- Architectural invariants (zero deps, React containment) are enforced by tests,
  not convention.

## 9. Docs and release

- Public API changes update the README in the same commit.
- Honest semver: breaking means major.
- Every limitation documented with its extension point.
- CHANGELOG entry per user-visible change.
- Experimental APIs are `unstable_`-prefixed and may change in a minor.

## 10. When rules conflict

1. Correctness and security
2. Composition (§0)
3. Developer experience — types, errors, defaults
4. Performance
5. Internal elegance

Internal elegance loses. A slightly repetitive primitive that composes cleanly
beats a clever abstraction that forces a configuration prop.
