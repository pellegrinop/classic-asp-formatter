# Classic ASP VS Code Extension 🚀

[![Visual Studio Marketplace](https://img.shields.io/badge/Marketplace-Classic%20ASP-blue?logo=visual-studio-code)](https://marketplace.visualstudio.com/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Advanced syntax highlighting, formatting, and diagnostics for **Classic ASP (.asp)** files. Built to bring a modern development experience to legacy codebase.

## ✨ Features

- 🎨 **Advanced Syntax Highlighting**: Full support for VBScript and HTML in the same file.
- 🧹 **Intelligent Formatter**:
  - Automatic indentation based on keywords (`If`, `For`, `Select Case`, `Function`, `Sub`, `Class`, `With`).
  - Automatic operator spacing (`=`, `&`, `+`, `<>`, etc.).
  - **String Preservation**: Never touches VBScript string contents (protects inline HTML).
  - **Empty Lines**: Preserves vertical whitespace while collapsing redundant blank lines (max 1).
- 🛠️ **Diagnostics**: Reports common block balancing errors for `If`, `Function`, `Sub`, `For`, `While`, and `Do` blocks.

## 🛠️ Installation & Usage

### Quick Installation
1. Download the latest `.vsix` file from the [Releases](https://github.com/pellegrinop/classic-asp-formatter/releases) section.
2. In VS Code, go to **Extensions** (`Ctrl+Shift+X`) -> `...` (More Actions) -> **Install from VSIX...**.
3. Ensure the language in the bottom right is set to **Classic ASP**.

### Commands
- `Shift + Alt + F`: Format Document.
- `Ctrl + Shift + P` -> `ASP: Check Extension Activity`: Verify if the extension is active.

## 🚀 Development & Testing

Built with **TypeScript** to ensure reliability and maintainability.

### Setup
```bash
npm install
```

### Build
```bash
# Compile TypeScript to JavaScript
npm run compile
```

### Automated Tests
A suite of unit tests ensures formatting logic remains robust:
```bash
npm test
```

### Packaging
```bash
# Lint the code
npm run lint

# Create .vsix package
npx vsce package
```

## 📄 License
Released under the MIT License. See [LICENSE.txt](LICENSE.txt) for details.

---
Built with ❤️ for ASP developers.
