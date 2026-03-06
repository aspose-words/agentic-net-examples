using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load a DOTM template that contains the desired styles
        string templatePath = @"C:\Templates\MyTemplate.dotm";
        Document doc = new Document(templatePath);

        // Enable the source‑styles switch so that styles from the template are applied automatically
        doc.AutomaticallyUpdateStyles = true;

        // Create a DocumentBuilder positioned at the end of the document (default position)
        DocumentBuilder builder = new DocumentBuilder(doc);

        // HTML fragment to be inserted dynamically
        string html = @"
            <h1>Report Title</h1>
            <p>This paragraph contains <b>bold</b> and <i>italic</i> text.</p>
            <ul>
                <li>First item</li>
                <li>Second item</li>
            </ul>";

        // Insert the HTML using builder formatting and remove the extra empty paragraph that InsertHtml adds
        builder.InsertHtml(html, HtmlInsertOptions.UseBuilderFormatting | HtmlInsertOptions.RemoveLastEmptyParagraph);

        // Save the resulting document
        string outputPath = @"C:\Output\GeneratedReport.docx";
        doc.Save(outputPath);
    }
}
