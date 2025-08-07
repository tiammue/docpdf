# Contributing to Document Converter

Thank you for your interest in contributing to Document Converter! This document provides guidelines for contributing to the project.

## Getting Started

1. Fork the repository
2. Clone your fork locally
3. Create a new branch for your feature or bug fix
4. Make your changes
5. Test your changes thoroughly
6. Submit a pull request

## Development Setup

### Prerequisites
- Python 3.7 or higher
- Microsoft Word (for testing DOC/DOCX to PDF conversion)

### Installation
```bash
git clone https://github.com/yourusername/document-converter.git
cd document-converter
pip install -r requirements.txt
```

### Running the Application
```bash
python document_converter.py
```

### Building the Executable
```bash
python -m PyInstaller --onefile --windowed --name "DocumentConverter" document_converter.py
```

## Code Style

- Follow PEP 8 Python style guidelines
- Use meaningful variable and function names
- Add comments for complex logic
- Keep functions focused and small
- Handle errors gracefully

## Testing

Before submitting a pull request:

1. Test both conversion directions (DOC/DOCX ↔ PDF)
2. Test with various file types and sizes
3. Verify the GUI remains responsive during conversion
4. Test the executable build process
5. Ensure error handling works correctly

## Reporting Issues

When reporting issues, please include:

- Operating system and version
- Python version
- Microsoft Word version (if applicable)
- Steps to reproduce the issue
- Expected vs actual behavior
- Any error messages or logs

## Feature Requests

We welcome feature requests! Please:

- Check if the feature already exists or is planned
- Describe the use case and benefits
- Consider implementation complexity
- Be open to discussion and feedback

## Pull Request Process

1. Update documentation if needed
2. Add or update tests for new functionality
3. Ensure all tests pass
4. Update CHANGELOG.md with your changes
5. Write clear commit messages
6. Keep pull requests focused on a single feature/fix

## Code of Conduct

- Be respectful and inclusive
- Focus on constructive feedback
- Help others learn and grow
- Maintain a positive environment

Thank you for contributing!