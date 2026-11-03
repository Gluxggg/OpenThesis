# Contributing to OpenThesis

Thanks for your interest! OpenThesis is an early-stage project aiming to be the universal document template engine.

## Ways to Contribute

### 🏫 Submit a Template
The most valuable contribution: **a real `.docx` template from your university or favorite journal.**

1. Put the `.docx` file in a new issue or PR
2. Tell us the institution name and document type (thesis / journal / official)
3. We'll add it to `examples/` and verify the parser handles it correctly

### 🐛 Report a Bug
Open an issue with:
- The `.docx` template that failed to parse (if applicable)
- The CLI command you ran
- The error output

### 💻 Code Contributions

```bash
git clone https://github.com/Gluxggg/OpenThesis.git
cd OpenThesis
pnpm install
pnpm check
```

**Project structure:**
- `packages/document-schema` — Type definitions (no deps)
- `packages/markdown-parser` — Markdown and front-matter import
- `packages/template-parser` — DOCX parsing logic
- `packages/docx-renderer` — DOCX generation
- `packages/equation-engine` — Equation rendering
- `packages/cli` — Command-line interface

**Before submitting a PR:**
- Run `pnpm check` and ensure the build and all tests pass
- Add tests for new functionality (if applicable)
- Follow the existing code style

### 📝 Improve Documentation
README, CLAUDE.md, inline comments — all improvements welcome.

## Development

```bash
pnpm install          # Install dependencies
pnpm build            # Build all packages
pnpm test             # Run the automated test suite
pnpm check            # Build, then run all tests (same gate as CI)
pnpm cli parse ...    # Test the parser
pnpm cli import ...   # Convert Markdown to structured JSON
pnpm cli build ...    # Test the renderer
pnpm cli init         # Create sample content
```

## License

By contributing, you agree that your contributions will be licensed under the MIT License.
