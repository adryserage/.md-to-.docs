# Markdown to DOCX Converter

A Python utility that converts Markdown files to Microsoft Word (DOCX) format. This tool can process both individual files and entire directories, making it perfect for batch converting documentation.

## Features

- Convert single Markdown files to DOCX
- Batch convert entire directories
- Preserves Markdown formatting including:
  - Headers
  - Lists (ordered and unordered)
  - Code blocks
  - Emphasis (bold, italic)
  - Links
  - Horizontal rules
- Simple command-line interface
- Error handling with informative messages
- Cross-platform compatibility

## Installation

1. Clone this repository:

```bash
git clone https://github.com/yourusername/.md-to-.docs.git
cd .md-to-.docs
```

2. Install the required dependencies:

```bash
pip install python-docx markdown beautifulsoup4
```

## Usage

The converter provides two main options:

1. Converting a single markdown file
2. Converting all markdown files in a directory

### Command Line Arguments

- `-i` or `--input`: Path to input markdown file or directory
- `-o` or `--output`: (Optional) Path to output file or directory
  - For single file: the output .docx file path
  - For directory: the output directory for converted files

### Examples

#### Converting a Single File

```bash
# Basic usage - output will be in the same directory as input
python md_to_docx_converter.py -i path/to/your/file.md

# Specify output file
python md_to_docx_converter.py -i path/to/your/file.md -o path/to/output.docx
```

#### Converting an Entire Directory

```bash
# Convert all .md files in a directory to a new directory
python md_to_docx_converter.py -i path/to/markdown/files -o path/to/output/directory
```

#### Real-world Example

```bash
# Convert all markdown files from docs/md to docs/word
python md_to_docx_converter.py -i ./docs/md -o ./docs/word
```

## Requirements

- Python 3.6 or higher
- Dependencies:
  - python-docx
  - markdown
  - beautifulsoup4

## Error Handling

The converter provides clear error messages for common issues:

- File not found
- Permission denied
- Invalid input/output paths
- Markdown parsing errors
- Word document creation errors

When batch converting files, the tool will:

- Continue processing remaining files if one fails
- Provide a summary of successful and failed conversions
- Display specific error messages for each failed conversion

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
