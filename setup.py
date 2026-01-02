"""
Enhanced setup.py for packaging the Marketing Dashboard Automator
"""

from setuptools import setup, find_packages
import os

# Read README for long description
with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

# Read version from version file
with open("VERSION", "r", encoding="utf-8") as fh:
    version = fh.read().strip()

setup(
    name="marketing-dashboard-automator",
    version=version,
    author="Marketing Analytics Team",
    author_email="analytics@company.com",
    description="Advanced PPT Data Extraction & Intelligent Dashboard Generation for Marketing Analytics",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/company/marketing-dashboard-automator",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Marketing",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Operating System :: OS Independent",
        "Operating System :: Microsoft :: Windows",
        "Operating System :: MacOS",
        "Operating System :: POSIX :: Linux",
    ],
    python_requires=">=3.8",
    install_requires=[
        "python-pptx>=1.0.0",
        "pandas>=2.0.0",
        "openpyxl>=3.0.0",
        "plotly>=5.0.0",
        "pyyaml>=6.0.0",
        "numpy>=1.24.0",
        "tkinterdnd2>=0.3.0;platform_system=='Windows'",
        "Pillow>=10.0.0",
    ],
    extras_require={
        "dev": [
            "black>=23.0.0",
            "flake8>=6.0.0",
            "pytest>=7.0.0",
            "pytest-cov>=4.0.0",
            "pre-commit>=3.0.0",
        ],
        "pdf": [
            "reportlab>=4.0.0",
            "python-docx>=0.8.11",
        ],
        "advanced": [
            "scikit-learn>=1.3.0",
            "seaborn>=0.12.0",
            "jinja2>=3.0.0",
            "requests>=2.31.0",
        ],
    },
    entry_points={
        "console_scripts": [
            "marketing-dashboard=main:main",
            "mda-process=cli:process",
            "mda-export=cli:export",
        ],
    },
    package_data={
        "": [
            "config/*.yaml",
            "templates/*.html",
            "assets/*.png",
            "assets/*.ico",
        ],
    },
    include_package_data=True,
    keywords=[
        "marketing",
        "dashboard",
        "analytics",
        "powerpoint",
        "excel",
        "reporting",
        "automation",
        "data-extraction",
    ],
    project_urls={
        "Bug Reports": "https://github.com/company/marketing-dashboard-automator/issues",
        "Source": "https://github.com/company/marketing-dashboard-automator",
        "Documentation": "https://github.com/company/marketing-dashboard-automator/wiki",
    },
)