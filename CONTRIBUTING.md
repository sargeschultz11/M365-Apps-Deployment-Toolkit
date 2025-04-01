# Contributing to M365 Apps Deployment Toolkit

Thank you for your interest in contributing to the M365 Apps Deployment Toolkit! This document provides guidelines and instructions for contributing to this project.

## Code of Conduct

By participating in this project, you are expected to uphold our [Code of Conduct](./CODE_OF_CONDUCT.md).

## How Can I Contribute?

### Reporting Bugs

If you encounter a bug, please create an issue using the bug report template and include:

- A clear and descriptive title
- Steps to reproduce the problem
- Expected behavior
- Actual behavior
- Screenshots if applicable
- Your environment details (OS, PowerShell version, etc.)

### Suggesting Enhancements

If you have ideas for enhancements:

- Use the feature request template
- Clearly describe the enhancement and its benefits
- Provide examples of how the enhancement would work

### Pull Requests

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Test your changes thoroughly
5. Commit your changes (`git commit -m 'Add some amazing feature'`)
6. Push to the branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

### Development Workflow

1. **Set up your environment**: Ensure you have PowerShell 5.1+ installed
2. **Testing**: Test all changes in multiple environments before submitting
3. **Documentation**: Update documentation to reflect your changes

## Styleguides

### Git Commit Messages

- Use the present tense ("Add feature" not "Added feature")
- Use the imperative mood ("Move cursor to..." not "Moves cursor to...")
- Limit the first line to 72 characters or less
- Reference issues and pull requests liberally after the first line

### PowerShell Styleguide

- Follow [PowerShell Best Practices](https://docs.microsoft.com/en-us/powershell/scripting/developer/cmdlet/strongly-encouraged-development-guidelines)
- Use proper indentation (4 spaces, not tabs)
- Include appropriate comments
- Use meaningful variable and function names
- Follow verb-noun naming for functions

### XML Configuration Files

- Use proper indentation
- Include comments explaining configuration options
- Maintain backward compatibility when possible

## Release Process

Our release process follows these steps:

1. Version updates in the script files
2. Changelog updates
3. Testing across multiple environments
4. Release creation with detailed notes

## Additional Notes

- Please focus on improving reliability, adding features, or enhancing documentation
- Consider backward compatibility when making changes
- Test thoroughly in multiple environments

Thank you for contributing to the M365 Apps Deployment Toolkit!
