# Contributing to Docx2MD Converter

We love your input! We want to make contributing to this project as easy and transparent as possible, whether it's:

- Reporting a bug
- Discussing the current state of the code
- Submitting a fix
- Proposing new features
- Becoming a maintainer

## Development Process

We use GitHub to host code, to track issues and feature requests, as well as accept pull requests.

## Pull Requests

Pull requests are the best way to propose changes to the codebase. We actively welcome your pull requests:

1. Fork the repo and create your branch from `main`.
2. If you've added code that should be tested, add tests.
3. If you've changed APIs, update the documentation.
4. Ensure the test suite passes.
5. Make sure your code lints.
6. Issue that pull request!

## Development Setup

### Prerequisites

- **VS Code**: Latest version
- **Node.js**: Version 16 or higher
- **Python**: Version 3.6 or higher
- **Git**: Latest version

### Setup Steps

1. **Clone the repository**:
   ```bash
   git clone https://github.com/your-username/docx2md-converter.git
   cd docx2md-converter
   ```

2. **Install dependencies**:
   ```bash
   npm install
   ```

3. **Install Python dependencies** (optional, for enhanced features):
   ```bash
   pip install python-docx olefile
   ```

4. **Open in VS Code**:
   ```bash
   code .
   ```

5. **Start development**:
   - Press `F5` to launch Extension Development Host
   - Make changes to the code
   - The extension will reload automatically

### Project Structure

```
docx2md-converter/
├── src/
│   ├── extension.ts          # Main extension entry point
│   └── test/
│       └── extension.test.ts # Extension tests
├── python/
│   └── converter.py          # Python conversion script
├── package.json              # Extension manifest
├── tsconfig.json            # TypeScript configuration
├── README.md                # Project documentation
├── CHANGELOG.md             # Version history
└── LICENSE                  # MIT license
```

## Coding Standards

### TypeScript/JavaScript

- Use TypeScript for all new code
- Follow the existing code style
- Use meaningful variable and function names
- Add JSDoc comments for public APIs
- Use `async/await` instead of callbacks where possible

### Python

- Follow PEP 8 style guidelines
- Use type hints where possible (Python 3.6+ compatible)
- Add docstrings to all functions and classes
- Use meaningful variable names
- Handle exceptions gracefully

### General Guidelines

- Keep functions small and focused
- Write self-documenting code
- Add comments for complex logic
- Use consistent naming conventions
- Prefer composition over inheritance

## Testing

### Running Tests

```bash
npm test
```

### Writing Tests

- Add tests for new functionality
- Update tests when changing existing code
- Use descriptive test names
- Test both success and error scenarios

### Manual Testing

1. Test with various document types:
   - Simple DOCX files
   - Complex DOCX with images and tables
   - DOC files
   - Corrupted files

2. Test different Python versions:
   - Minimum supported (3.6)
   - Latest stable version

3. Test on different platforms:
   - Windows
   - macOS
   - Linux

## Submitting Issues

### Bug Reports

When filing an issue, make sure to answer these questions:

1. What version of VS Code are you using?
2. What version of Python are you using?
3. What operating system are you using?
4. What did you do?
5. What did you expect to see?
6. What did you see instead?

### Feature Requests

We welcome feature requests! Please:

1. Check if the feature already exists
2. Check if someone else has requested it
3. Describe the feature in detail
4. Explain why it would be useful
5. Provide examples if possible

## Code Review Process

1. All submissions require review
2. We use GitHub pull requests for this purpose
3. Maintainers will review and provide feedback
4. Address feedback and update your PR
5. Once approved, we'll merge your changes

## Community Guidelines

### Our Pledge

We pledge to make participation in our project a harassment-free experience for everyone, regardless of age, body size, disability, ethnicity, gender identity and expression, level of experience, nationality, personal appearance, race, religion, or sexual identity and orientation.

### Our Standards

Examples of behavior that contributes to creating a positive environment include:

- Using welcoming and inclusive language
- Being respectful of differing viewpoints and experiences
- Gracefully accepting constructive criticism
- Focusing on what is best for the community
- Showing empathy towards other community members

### Enforcement

Project maintainers are responsible for clarifying the standards of acceptable behavior and are expected to take appropriate and fair corrective action in response to any instances of unacceptable behavior.

## Recognition

Contributors will be recognized in:

- README.md contributors section
- CHANGELOG.md for significant contributions
- Release notes for major features

## Questions?

Don't hesitate to ask questions! You can:

- Open an issue for discussion
- Start a GitHub Discussion
- Contact the maintainers directly

## License

By contributing, you agree that your contributions will be licensed under the MIT License.

---

Thank you for contributing to Docx2MD Converter! 🎉
