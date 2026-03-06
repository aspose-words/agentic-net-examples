using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertHtmlIntoDotm
{
    static void Main()
    {
        // Path to the DOTM template that contains a bookmark named "HtmlContent".
        const string templatePath = @"C:\Templates\ReportTemplate.dotm";

        // Load the DOTM template.
        Document doc = new Document(templatePath);

        // Create a DocumentBuilder for the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // HTML fragment to be inserted dynamically.
        const string html = @"
            <h2 style='color:#2E86C1;'>Sales Summary</h2>
            <table border='1' cellpadding='5' cellspacing='0'>
                <tr><th>Product</th><th>Units Sold</th><th>Revenue</th></tr>
                <tr><td>Widget A</td><td>120</td><td>$3,600</td></tr>
                <tr><td>Widget B</td><td>85</td><td>$2,550</td></tr>
            </table>";

        // Move the cursor to the bookmark where the HTML should be placed.
        builder.MoveToBookmark("HtmlContent");

        // Insert the HTML. Use builder formatting as base formatting.
        builder.InsertHtml(html, HtmlInsertOptions.UseBuilderFormatting);

        // Save the resulting document. The format is inferred from the file extension.
        const string outputPath = @"C:\Output\ReportWithHtml.docx";
        doc.Save(outputPath);
    }
}
