using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class HtmlInserter
{
    /// <summary>
    /// Inserts the specified HTML into a Word template at a merge field named "HTML".
    /// The <paramref name="sourceStyles"/> flag determines whether the formatting of the
    /// DocumentBuilder is applied to the inserted HTML.
    /// </summary>
    /// <param name="templatePath">Path to the DOCX template file.</param>
    /// <param name="outputPath">Path where the resulting document will be saved.</param>
    /// <param name="html">HTML string to insert.</param>
    /// <param name="sourceStyles">If true, use builder formatting (HtmlInsertOptions.UseBuilderFormatting); otherwise insert HTML with its own formatting.</param>
    public static void InsertHtmlFromTemplate(string templatePath, string outputPath, string html, bool sourceStyles)
    {
        // Load the template document.
        Document doc = new Document(templatePath);

        // Create a DocumentBuilder for the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the merge field where the HTML should be placed.
        // The template must contain a MERGEFIELD named "HTML".
        builder.MoveToMergeField("HTML");

        // Choose insertion options based on the sourceStyles flag.
        HtmlInsertOptions options = sourceStyles
            ? HtmlInsertOptions.UseBuilderFormatting
            : HtmlInsertOptions.None;

        // Insert the HTML with the selected options.
        builder.InsertHtml(html, options);

        // Save the modified document.
        doc.Save(outputPath);
    }

    // Example usage.
    public static void Main()
    {
        string template = @"C:\Templates\ReportTemplate.docx";
        string output = @"C:\Results\ReportResult.docx";

        // Sample HTML fragment.
        string html = @"
            <h2 align='center'>Quarterly Summary</h2>
            <p>This quarter's <b>revenue</b> increased by <span style='color:green'>15%</span>.</p>";

        // Insert HTML using source styles (builder formatting).
        InsertHtmlFromTemplate(template, output, html, sourceStyles: true);
    }
}
