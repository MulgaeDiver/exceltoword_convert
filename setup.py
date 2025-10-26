from setuptools import setup, find_packages

setup(
    name="excel-to-word-converter",
    version="1.0.0",
    description="Excel 파일을 Word 문서로 변환하는 자동화 프로그램",
    author="Your Name",
    author_email="your.email@example.com",
    packages=find_packages(),
    install_requires=[
        "pandas>=2.2.0",
        "openpyxl>=3.1.2",
        "python-docx>=1.1.0",
        "streamlit>=1.28.1",
    ],
    entry_points={
        "console_scripts": [
            "excel-to-word=run_app:main",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.13",
    ],
    python_requires=">=3.8",
)


