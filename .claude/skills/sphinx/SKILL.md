---
name: sphinx
description: Set up and work with Sphinx (npm package `sphinx-understand`, github.com/parthh01/sphinx), a comprehension checkpoint that pauses `git push` to protected branches until the human answers questions about the change. Use when the user asks to install, configure, check, pause or uninstall Sphinx, mentions `.sphinxconfig`/`.sphinxignore`/`.sphinxcontext.md`, or when a `git push` prints "Sphinx: understanding check needed" or a Sphinx questionnaire URL.
---

# Sphinx

Sphinx is a CLI, not a library. It installs machine-wide Git hooks (`core.hooksPath`) and, for Claude Code, hooks plus a `sphinx` MCP server. On a matching push, Git waits while a separate reviewer model asks the human questions in a local page (`http://127.0.0.1:<port>/…`). Once they pass and choose **Continue push**, the same push goes through.

This repository already has its policy files committed: `.sphinxconfig` (pushes to `main`/`master`, standard depth, mixed distance), `.sphinxignore`, `.sphinxcontext.md`.

## Install (on the user's own machine)

Requires Node.js >= 22.19, Git, macOS/Linux/WSL. Not useful in an ephemeral cloud container: the questionnaire is on loopback and the hooks vanish with the container.

```sh
npm install -g sphinx-understand
sphinx install --yes --preset claude-economy --agents claude   # or plain `sphinx install` for guided setup
sphinx doctor                                                   # must end with "Configured integrations are in place."
```

Then start a new Claude Code session; `/mcp` should list `sphinx` as connected. `sphinx init` in a repo creates the three policy files without overwriting existing ones.

`sphinx install` changes global state (Git `core.hooksPath`, `~/.claude/settings.json`): run it only when the user asks. Undo with `sphinx uninstall`.

## When a push is blocked

- Git prints a questionnaire URL and waits up to 25 minutes. Tell the user to open it and answer. You cannot answer for them: answers go straight to Sphinx and you only see status.
- Never start a second copy of a push that is still waiting. If your shell timed out and Git exited, run the same push once more; a saved pass is reused for 24h unless code, destination or policy changed.
- No browser (SSH/headless): the user can run `sphinx check` in a terminal.
- Do not skip, `sphinx yolo`, or `--no-verify` unless the user says to. Skips are logged.

## Commands

| Command | Use |
| --- | --- |
| `sphinx init` | Create repo policy files |
| `sphinx doctor` | Check hook and integration coverage |
| `sphinx status` | Policy, pending checks, recent skips |
| `sphinx check` | Run or resume a check in the terminal (`--staged` for commits) |
| `sphinx pending` | List unfinished checks |
| `sphinx yolo [--reason "…"]` | Toggle checks off/on (repo, or global outside a repo) |
| `sphinx config` | Reviewer model, limits, `--distance low\|mixed\|high`, `--interface terminal\|browser`, `--max-model-calls N` |
| `sphinx auth` | Save an OpenRouter key or OpenAI-compatible endpoint |
| `sphinx uninstall` | Remove integrations, keep unrelated settings |

Every command takes `--help`; `--json` gives machine-readable output.

## Privacy and cost

The reviewer provider receives changed code, related files, answers and `.sphinxcontext.md`. `.env`, private keys and common credential files are always excluded; add other sensitive paths to `.sphinxignore` before the first check. A check makes about 5-8 model calls plus about one per answer, with no cap unless set (`sphinx config --max-model-calls`).
