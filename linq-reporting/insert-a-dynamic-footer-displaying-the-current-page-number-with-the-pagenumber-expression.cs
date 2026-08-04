using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some content to generate multiple pages.
        builder.Writeln("First page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Third page content.");

        // Move the cursor to the primary footer and insert the dynamic page number tag.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("{=PageNumber}");

        // Build the report (no data source is required for the footer tag).
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, new object());

        // Save the resulting document.
        doc.Save("DynamicFooter.docx");
    }
}
