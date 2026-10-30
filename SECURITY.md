# Security policy

XSheet opens `.xlsx` files that you did not create. It is built to treat every file as untrusted input, so a crash, hang or escape caused by a crafted file is a security issue.

## Supported versions

| Version | Supported |
| --- | --- |
| Latest release (0.3.x) | ✅ |
| Older releases | ❌ — please upgrade first |

## Reporting a vulnerability

Please **do not open a public issue** for a vulnerability.

Use GitHub's private reporting instead: **Security → Report a vulnerability** on this repository. Include:

- the XSheet version and how you load it (`view`, `renderWorkbook`, the standalone build, …);
- what happens (hang, crash, memory growth, script execution, network access, …);
- a file that reproduces it, or a description of how to build one. If you send a file, make it from scratch — please do not send real documents that contain other people's data.

This is a personal project, so replies may take a few days. Once a fix is released, the report is credited in the release notes unless you ask otherwise.

## What counts

In scope — anything that breaks these guarantees:

- **No network.** The library never fetches or sends anything by itself.
- **No code execution.** Nothing from the file is evaluated, and no HTML from the file reaches the DOM (text goes in through `textContent` only).
- **Bounded work.** Reading stops at the limits in `DEFAULT_LIMITS` (unzipped size, part sizes, counts, shape formulas, images, chart points). A file that makes XSheet hang, use unbounded memory, or bypass a limit is in scope.
- **No content in errors.** Exception messages never contain cell values or sheet names.

Out of scope:

- Raising `limits` yourself and then running out of memory.
- Rendering differences from Excel — please open a normal issue for those.
- Vulnerabilities in your own page (for example, a CSP that allows `unsafe-eval` for other reasons).
