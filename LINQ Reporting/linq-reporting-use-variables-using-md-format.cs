using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class LinqReportingWithMarkdown
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a placeholder that will be replaced by markdown text.
        builder.Writeln("{{ReportContent}}");

        // Sample markdown string – this could come from any LINQ query or data source.
        string markdownText = "# Sales Report\r\n" +
                              "## Summary\r\n" +
                              "- **Total Orders:** 124\r\n" +
                              "- **Revenue:** $12,340.00\r\n" +
                              "\r\n" +
                              "### Details\r\n" +
                              "| Product | Qty | Price |\r\n" +
                              "|---------|-----|-------|\r\n" +
                              "| Widget  |  10 | $5.00 |\r\n" +
                              "| Gizmo   |   5 | $7.50 |";

        // Configure find‑replace options to treat the replacement string as Markdown.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacementFormat = ReplacementFormat.Markdown;

        // Replace the placeholder with the markdown text using the configured options.
        doc.Range.Replace("{{ReportContent}}", markdownText, options);

        // Save the resulting document.
        doc.Save("LinqReporting_Markdown.docx");
    }
}
