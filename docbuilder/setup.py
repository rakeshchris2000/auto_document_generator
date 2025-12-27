"""
Setup script for Google Docs API Client Library
"""

from setuptools import setup, find_packages
import os

# Read the README file for long description
def read_readme():
    readme_path = os.path.join(os.path.dirname(__file__), 'README.md')
    if os.path.exists(readme_path):
        with open(readme_path, 'r', encoding='utf-8') as f:
            return f.read()
    return "A comprehensive Python client library for the Google Docs API"

# Read requirements
def read_requirements():
    req_path = os.path.join(os.path.dirname(__file__), 'requirements.txt')
    requirements = []
    if os.path.exists(req_path):
        with open(req_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    requirements.append(line)
    return requirements

setup(
    name="google-docs-client",
    version="1.0.0",
    author="Google Docs Client Library",
    author_email="support@example.com",
    description="A comprehensive Python client library for the Google Docs API",
    long_description=read_readme(),
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/google-docs-client",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "Topic :: Office/Business :: Office Suites",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
    install_requires=read_requirements(),
    keywords="google docs api client library automation document",
    project_urls={
        "Bug Reports": "https://github.com/yourusername/google-docs-client/issues",
        "Source": "https://github.com/yourusername/google-docs-client",
        "Documentation": "https://github.com/yourusername/google-docs-client#readme",
    },
    entry_points={
        'console_scripts': [
            'google-docs-demo=google_docs_client.main:main',
        ],
    },
    include_package_data=True,
    zip_safe=False,
)
