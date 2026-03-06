using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class InsertHtmlDynamic
{
    static void Main()
    {
        // Path to the WORDML (DOCX) template that contains a merge field named "HTML_CONTENT"
        const string templatePath = @"C:\Templates\WordmlTemplate.docx";

        // Load the template document (lifecycle rule: load)
        Document doc = new Document(templatePath);

        // Create a DocumentBuilder for editing the document (lifecycle rule: create)
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the merge field that will be replaced by HTML.
        // The template should contain <<MERGEFIELD HTML_CONTENT \b Content>> or a bookmark with the same name.
        builder.MoveToMergeField("HTML_CONTENT");

        // The HTML string to insert dynamically.
        string html = @"
            <h2 style='color:#2E86C1;'>Dynamic Report Title</h2>
            <p>This paragraph contains <b>bold</b>, <i>italic</i> and a <a href='https://www.example.com'>link</a>.</p>
            <ul>
                <li>First item</li>
                <li>Second item</li>
            </ul>";

        // Insert the HTML into the document.
        // Use builder formatting as base and remove the extra empty paragraph that Word adds after block‑level HTML.
        builder.InsertHtml(html, HtmlInsertOptions.UseBuilderFormatting | HtmlInsertOptions.RemoveLastEmptyParagraph);

        // Save the populated document (lifecycle rule: save)
        const string outputPath = @"C:\Output\DynamicReport.docx";
        doc.Save(outputPath);
    }
}
