# HTML to DOCX Converter

A Python library to convert HTML content into a .docx file, with a focus on handling complex tables, nested lists, and other common HTML elements with improved formatting.

## Features

- **Tables**: Handles complex tables, including those with `rowspan` and `colspan`, and tables with `<thead>`, `<tbody>`, and `<tfoot>` sections.
- **Lists**: Supports nested ordered (`<ol>`) and unordered (`<ul>`) lists with proper indentation and styling for different levels.
- **Text Formatting**: Converts common inline text formatting tags like `<b>`, `<strong>`, `<i>`, `<em>`, `<u>`, `<s>`, `<code>`.
- **Headings**: Converts `<h1>` to `<h6>` tags to corresponding heading styles in the docx file.
- **Blockquotes**: Formats `<blockquote>` tags with indentation and a background color.
- **Pre-formatted Text**: Renders `<pre>` blocks with a monospace font and a background color, preserving whitespace.
- **Images**: Handles `<img>` tags with both URL and local file paths.
- **Hyperlinks**: Converts `<a>` tags to clickable hyperlinks.
- **Horizontal Rules**: Renders `<hr>` tags as horizontal lines.

## Installation

To install the library, you can clone the repository and install it using pip:

```bash
git clone https://github.com/your-username/html2docx.git
cd html2docx
pip install .
```

## Usage

Here are some examples of how to use the library.

### Converting an HTML String

You can convert a string of HTML content directly into a new docx document.

```python
from htmldocx import HtmlToDocx
from docx import Document

html_content = """
    <h1>My Document</h1>
    <p>This is a paragraph.</p>
    <ul>
        <li>Item 1</li>
        <li>Item 2
            <ol>
                <li>Nested item 2.1</li>
            </ol>
        </li>
    </ul>
"""

document = Document()
parser = HtmlToDocx()
parser.add_html_to_document(html_content, document)
document.save("my_document.docx")
```

### Inserting HTML into an Existing Document

The library can also be used to insert HTML content into an existing document at a specific location. The following example shows how to replace a placeholder paragraph with HTML content.

```python
from htmldocx import HtmlToDocx
from docx import Document
from copy import deepcopy

def insert_html_at_placeholder(doc, placeholder, html):
    # Find the placeholder paragraph
    for p in doc.paragraphs:
        if placeholder in p.text:
            # Create a temporary document to parse the HTML
            temp_doc = Document()
            parser = HtmlToDocx()
            parser.add_html_to_document(html, temp_doc)

            # Get the parent element and index of the placeholder paragraph
            parent = p._element.getparent()
            idx = parent.index(p._element)

            # Insert the parsed elements from the temporary document
            for element in temp_doc._body._element:
                parent.insert(idx, deepcopy(element))
                idx += 1
            
            # Remove the placeholder paragraph
            parent.remove(p._element)
            return

# --- Example usage ---
doc = Document("template.docx") # Your template with placeholder
html_to_insert = "<h2>Content to Insert</h2><p>This is some new content.</p>"
insert_html_at_placeholder(doc, "{{PLACEHOLDER}}", html_to_insert)
doc.save("final_document.docx")
```

## Contributing

Contributions are welcome! If you find a bug or have a feature request, please open an issue on GitHub. You can also fork the repository and submit a pull request.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.