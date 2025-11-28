"""
Automated PyInstaller Build Script
Builds all apps, cleans up build artifacts, and organizes output
"""
import os
import shutil
import subprocess
from pathlib import Path

# Configuration
APPS = [
    {"name": "DrillDownApp", "script": "drill_down_app.py"},
    {"name": "PivotAnalysisApp", "script": "pivot_analysis_app.py"},
    {"name": "StockAnalysisApp", "script": "stock_analysis_app.py"},
]

DIST_PATH = "Portfolio_analyzers"
CLEANUP_DIRS = ["build", "__pycache__"]
CLEANUP_PATTERNS = ["*.spec"]


def clean_artifacts():
    """Remove build artifacts and old spec files"""
    print("🧹 Cleaning build artifacts...")
    
    # Remove directories
    for dir_name in CLEANUP_DIRS:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"  ✓ Removed {dir_name}/")
    
    # Remove spec files
    for pattern in CLEANUP_PATTERNS:
        for file in Path(".").glob(pattern):
            file.unlink()
            print(f"  ✓ Removed {file}")
    
    # Clean Utils pycache
    utils_pycache = Path("Utils/__pycache__")
    if utils_pycache.exists():
        shutil.rmtree(utils_pycache)
        print(f"  ✓ Removed Utils/__pycache__/")


def build_app(app_config):
    """Build a single app using PyInstaller"""
    name = app_config["name"]
    script = app_config["script"]
    
    print(f"\n🔨 Building {name}...")
    
    cmd = [
        "pyinstaller",
        "--noconsole",
        "--windowed",
        "--onefile",
        "--name", name,
        "--add-data", "Utils;Utils",
        "--distpath", DIST_PATH,
        script
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode == 0:
        print(f"  ✅ {name}.exe created successfully")
        return True
    else:
        print(f"  ❌ Failed to build {name}")
        print(result.stderr)
        return False


def copy_excels_folder():
    """Copy Excels folder to distribution directory"""
    print("\n📁 Copying Excels folder...")
    
    source = Path("Excels")
    dest = Path(DIST_PATH) / "Excels"
    
    if source.exists():
        if dest.exists():
            shutil.rmtree(dest)
        shutil.copytree(source, dest)
        print(f"  ✓ Copied Excels/ to {DIST_PATH}/Excels/")
    else:
        print("  ⚠️  Excels folder not found")


def create_readme():
    """Create a README in the distribution folder"""
    readme_content = """# Stock Navigation Automation - Portfolio Analyzers

## Applications

- **DrillDownApp.exe**: Drill-down analysis tool
- **PivotAnalysisApp.exe**: Pivot table analysis
- **StockAnalysisApp.exe**: Stock analysis tool

## Data Files

All Excel data files are located in the `Excels/` folder.

## Usage

Simply double-click any .exe file to launch the application.
No Python installation required.

---
Built with PyInstaller
"""
    
    readme_path = Path(DIST_PATH) / "README.txt"
    readme_path.write_text(readme_content)
    print(f"  ✓ Created README.txt")


def main():
    """Main build process"""
    print("=" * 60)
    print("Stock Navigation Automation - Build Script")
    print("=" * 60)
    
    # Step 1: Clean old artifacts
    clean_artifacts()
    
    # Step 2: Build all apps
    success_count = 0
    for app in APPS:
        if build_app(app):
            success_count += 1
    
    # Step 3: Copy Excels folder
    copy_excels_folder()
    
    # Step 4: Create README
    create_readme()
    
    # Step 5: Clean up build artifacts again
    print("\n🧹 Final cleanup...")
    clean_artifacts()
    
    # Summary
    print("\n" + "=" * 60)
    print(f"✨ Build complete: {success_count}/{len(APPS)} apps built successfully")
    print(f"📦 Output location: {DIST_PATH}/")
    print("=" * 60)


if __name__ == "__main__":
    main()