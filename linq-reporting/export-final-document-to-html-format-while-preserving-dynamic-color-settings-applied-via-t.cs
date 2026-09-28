using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare working directories.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "Template.docx");
        string outputPath = Path.Combine(workDir, "Report.html");

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a paragraph with dynamic text color and background color.
        // The colors are taken from the model properties.
        builder.Writeln("<<textColor [model.Color]>>");
        builder.Writeln("<<backColor [model.BackColor]>>Dynamic Colored Text<</backColor>>");
        builder.Writeln("<</textColor>>");

        // Insert an HTML snippet to demonstrate HTML export.
        builder.Writeln("<<[model.HtmlSnippet] -html>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Sample data model with color properties.
        ReportModel model = new ReportModel
        {
            Color = "Blue",
            BackColor = "LightGray",
            HtmlSnippet = "<b>Bold HTML content</b>"
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 3. Export the final document to HTML, preserving colors.
        // -----------------------------------------------------------------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        reportDoc.Save(outputPath, htmlOptions);
    }
}

// ---------------------------------------------------------------------
// Data model used by the template.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Text color name or HTML color code.
    public string Color { get; set; } = "Black";

    // Background color name or HTML color code.
    public string BackColor { get; set; } = "White";

    // HTML snippet to be inserted as raw HTML.
    public string HtmlSnippet { get; set; } = string.Empty;
}
