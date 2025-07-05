# 🤝 CONTRIBUTING

Thank you for considering contributing to **Order Checker**! 

## 🎯 How to Contribute

### 🐛 Reporting Bugs
1. **Check existing issues** first to avoid duplicates
2. **Create detailed bug report** with:
   - Steps to reproduce
   - Expected vs actual behavior
   - Screenshots/error messages
   - Operating system and Python version
   - Config file structure (remove sensitive data)

### ✨ Suggesting Features
1. **Open a feature request** issue
2. **Describe the use case** and benefits
3. **Provide examples** if possible
4. **Consider implementation complexity**

### 🔧 Code Contributions

#### Prerequisites
- Python 3.8+
- Basic knowledge of Playwright
- Understanding of Excel file manipulation
- Familiarity with web scraping concepts

#### Development Setup
```bash
# Fork and clone the repository
git clone https://github.com/your-username/order-checker.git
cd order-checker

# Set up virtual environment
python -m venv myenv
myenv\Scripts\activate  # Windows
source myenv/bin/activate  # Linux/Mac

# Install dependencies
pip install -r requirements.txt
python -m playwright install chromium
python setup_chromium.py
```

#### Making Changes
1. **Create a feature branch**: `git checkout -b feature/your-feature-name`
2. **Make your changes** following code style guidelines
3. **Test thoroughly** with different scenarios
4. **Update documentation** if needed
5. **Commit with clear messages**: `git commit -m "Add: feature description"`
6. **Push and create Pull Request**

## 📝 Code Style Guidelines

### Python Code Standards
- **PEP 8** compliance for code formatting
- **Type hints** for function parameters and returns
- **Docstrings** for classes and complex functions
- **Error handling** with try/catch blocks
- **Clean console output** with appropriate logging levels

### Example Code Structure
```python
def process_excel_file(file_path: str, sheet_name: str) -> bool:
    """
    Process Excel file with specific formatting rules.
    
    Args:
        file_path (str): Path to Excel file
        sheet_name (str): Name of the sheet to process
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Implementation here
        return True
    except Exception as e:
        print(f"❌ Error processing {sheet_name}: {str(e)}")
        return False
```

### Configuration Guidelines
- **JSON format** for all config files
- **Template files** for default configurations
- **Validation** before using config values
- **Clear error messages** for invalid configs
- **Security considerations** for sensitive data

## 🧪 Testing Guidelines

### Before Submitting
- [ ] **Test with different websites** (if applicable)
- [ ] **Verify Excel processing** with various file formats
- [ ] **Check executable packaging** works correctly
- [ ] **Validate error handling** with invalid inputs
- [ ] **Ensure console output** is clean and informative

### Test Scenarios
1. **Valid config with working credentials**
2. **Invalid config (missing fields, wrong selectors)**
3. **Network connectivity issues**
4. **Excel files with different structures**
5. **Edge cases (empty reports, special characters)**

## 📚 Documentation Updates

### Required Documentation
- Update **README.md** for new features
- Modify **CONFIG_GUIDE.md** for new selectors
- Add entries to **CHANGELOG.md**
- Include **inline comments** for complex logic

### Documentation Style
- **Clear headings** with emoji icons
- **Code examples** with proper syntax highlighting
- **Step-by-step instructions** for setup
- **Troubleshooting sections** for common issues
- **Screenshots** when helpful

## 🔒 Security Considerations

### Handling Sensitive Data
- **Never commit** actual credentials to repository
- **Use template files** for config examples
- **Add security warnings** in documentation
- **Sanitize logs** to remove sensitive information

### Code Security
- **Validate inputs** before processing
- **Escape user data** in web interactions
- **Handle errors gracefully** without exposing internals
- **Use secure connection methods**

## 🚀 Release Process

### Version Numbering
- **Major**: Breaking changes or significant new features
- **Minor**: New features that are backward compatible
- **Patch**: Bug fixes and small improvements

### Release Checklist
- [ ] Update version numbers in code
- [ ] Update CHANGELOG.md with changes
- [ ] Test executable packaging
- [ ] Verify all documentation is current
- [ ] Create release notes
- [ ] Tag the release in Git

## 💬 Communication

### Getting Help
- **GitHub Issues** for bugs and feature requests
- **GitHub Discussions** for questions and ideas
- **Email** for security-related concerns

### Code Review Process
1. **Automated checks** must pass
2. **Manual review** by maintainers
3. **Testing verification** in different environments
4. **Documentation review** for completeness

## 🙏 Recognition

### Contributors
All contributors will be recognized in:
- **README.md** contributors section
- **Release notes** for their specific contributions
- **GitHub contributors** page

### Types of Contributions
- **Code** improvements and new features
- **Documentation** updates and clarifications
- **Bug reports** with detailed information
- **Testing** and validation efforts
- **Design** improvements for UI/UX

---

## 📋 Quick Contribution Checklist

Before submitting your contribution:

- [ ] **Code follows** style guidelines
- [ ] **Tests pass** in multiple scenarios  
- [ ] **Documentation updated** appropriately
- [ ] **Commit messages** are clear and descriptive
- [ ] **No sensitive data** in commits
- [ ] **Feature branch** used for development
- [ ] **Pull request** has detailed description

Thank you for contributing to Order Checker! 🎉

---

*For questions about contributing, please open an issue or start a discussion.*
