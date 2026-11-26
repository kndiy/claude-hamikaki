# CLAUDE.md change log archive

Older change-log entries moved out of `../CLAUDE.md` on 2026-11-26, verbatim and in their original order. The newest entries stay in `CLAUDE.md`; when an entry leaves it, append it here.

- 2026-09-27: `M)` now accepts a claude.ai export .zip or folder and uses separate export rules. Tested on both ports with a fake export: extraction, detection in a subfolder, rules chosen, temp deleted. **Not yet tested with a real export.**
- 2026-09-27: Added `M) migrate`. Moved the shared launch code into `Invoke-ClaudeIn` / `claude_in`. Fixed the Windows stdout-capture bug, which also affected `S) settings`. Tested in a scratch tree on Windows and under Git Bash with a stub `claude`: bad path, source containing the target, settings, migrate, new-project detection → open. A real interactive session hasn't been tried yet.
- 2026-09-27: Added the `S) settings` option (Claude Code at the root, subRoot or group, with access to `_launcher`). Tested on Windows and under Git Bash with a stub `claude` that prints its cwd and args. A real interactive session hasn't been tried yet.
- 2026-09-27: Fixes found in testing: path-string navigation broke on 8.3 names (replaced with a stack); a crash on EOF input (added `Read-Line`); `wsl.exe` stderr threw under `Stop`; checks written as PowerShell cmdlets could not fail (now PowerShell errors count as failure).
