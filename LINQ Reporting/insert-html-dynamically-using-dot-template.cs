using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class InsertHtmlWithDotTemplate
{
    static void Main()
    {
        // Load the DOT (template) document that contains a bookmark named "HtmlPlaceholder".
        Document template = new Document("Template.dotx");

        // Create a data source object with a property that holds the HTML to insert.
        var data = new
        {
            HtmlContent = @"
                <h2 style='color:#2E86C1;'>Dynamic Title</h2>
                <p>This paragraph is <b>bold</b> and <i>italic</i>.</p>
                <ul>
                    <li>Item 1</li>
                    <li>Item 2</li>
                </ul>"
        };

        // Use ReportingEngine to populate the template. The template should contain a MERGEFIELD
        // like: MERGEFIELD HtmlContent \b Content
        // The "\b Content" switch tells Aspose.Words to treat the field value as HTML.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, data, "src");

        // If the template does not use a MERGEFIELD, you can also insert the HTML manually.
        // Move the cursor to the bookmark and insert the HTML string.
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.MoveToBookmark("HtmlPlaceholder");
        builder.InsertHtml(data.HtmlContent, HtmlInsertOptions.UseBuilderFormatting | HtmlInsertOptions.RemoveLastEmptyParagraph);

        // Save the resulting document.
        template.Save("Result.docx");
    }
}
