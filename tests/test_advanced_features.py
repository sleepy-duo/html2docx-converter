from htmldocx import HtmlToDocx
from docx import Document

def test_new_features_smoke_test():
    """A smoke test to check that the new features run without errors."""
    html_content = """
    <h1>Feature Test</h1>

    <h2>Blockquote</h2>
    <blockquote>
      <p>This is a blockquote. It should be indented and have a background color.</p>
      <p>This is the second paragraph in the blockquote.</p>
    </blockquote>

    <h2>Pre-formatted text</h2>
    <pre><code>
    def hello_world():
        print("Hello, world!")
    </code></pre>

    <h2>Inline elements</h2>
    <p>This is a paragraph with <b>bold</b>, <i>italic</i>, <u>underline</u>, <s>strikethrough</s>, and <code>inline code</code>.</p>

    <h2>Horizontal Rule</h2>
    <hr>

    <p>End of test.</p>
    """

    document = Document()
    parser = HtmlToDocx()
    parser.add_html_to_document(html_content, document)
    document.save("test_features_output.docx")
