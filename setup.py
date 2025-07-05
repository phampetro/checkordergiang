#!/usr/bin/env python3
"""
🚀 QUICK SETUP SCRIPT for Order Checker
Run this script to automatically set up the environment
"""

import os
import sys
import subprocess
from pathlib import Path

def run_command(command, description):
    """Run a command and handle errors"""
    print(f"🔄 {description}...")
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        print(f"✅ {description} completed!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {description} failed: {e}")
        print(f"Error output: {e.stderr}")
        return False

def main():
    print("🎯 ORDER CHECKER - QUICK SETUP")
    print("=" * 50)
    
    # Check Python version
    if sys.version_info < (3, 8):
        print("❌ Python 3.8+ is required!")
        return False
    
    print(f"✅ Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro} detected")
    
    # Check if virtual environment exists
    venv_path = Path("myenv")
    if not venv_path.exists():
        print("📦 Creating virtual environment...")
        if not run_command("python -m venv myenv", "Virtual environment creation"):
            return False
    else:
        print("✅ Virtual environment already exists")
    
    # Activate virtual environment and install dependencies
    if os.name == 'nt':  # Windows
        activate_cmd = "myenv\\Scripts\\activate"
        pip_cmd = "myenv\\Scripts\\pip"
        python_cmd = "myenv\\Scripts\\python"
    else:  # Linux/Mac
        activate_cmd = "source myenv/bin/activate"
        pip_cmd = "myenv/bin/pip" 
        python_cmd = "myenv/bin/python"
    
    # Install requirements
    if not run_command(f"{pip_cmd} install -r requirements.txt", "Installing Python packages"):
        return False
    
    # Install Playwright browsers
    if not run_command(f"{python_cmd} -m playwright install chromium", "Installing Chromium browser"):
        return False
    
    # Setup Chromium for packaging
    if Path("setup_chromium.py").exists():
        if not run_command(f"{python_cmd} setup_chromium.py", "Setting up Chromium for packaging"):
            print("⚠️  Chromium setup failed, but continuing...")
    
    # Create input/output directories
    Path("input").mkdir(exist_ok=True)
    Path("output").mkdir(exist_ok=True)
    print("✅ Created input/output directories")
    
    # Check if config exists
    config_path = Path("input/config.json")
    if not config_path.exists():
        print("📝 Config file not found - will be created on first run")
    else:
        print("✅ Config file found")
    
    print("\n" + "=" * 50)
    print("🎉 SETUP COMPLETED SUCCESSFULLY!")
    print("\n📋 Next steps:")
    print("1. Edit input/template.xlsx with your report list")
    print("2. Run the program: python check_oder.py")
    print("3. Update input/config.json with your credentials")
    print("4. Run again to start automation")
    print("\n📚 For detailed instructions, see README.md")
    
    return True

if __name__ == "__main__":
    try:
        success = main()
        if not success:
            print("\n❌ Setup failed! Please check the errors above.")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n⚠️  Setup interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        sys.exit(1)
