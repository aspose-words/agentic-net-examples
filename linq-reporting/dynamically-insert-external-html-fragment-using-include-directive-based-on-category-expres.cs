using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the working directory exists.
        string workDir = Directory.GetCurrentDirectory();

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(workDir, "Template.docx");
        string reportPath = Path.Combine(workDir, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a title.
        builder.Writeln("Dynamic HTML inclusion based on Category");
        builder.Writeln();

        // Conditional inclusion of HTML fragments.
        // If Category == "A" include HtmlA, else if Category == "B" include HtmlB.
        builder.Writeln("<<if [model.Category == \"A\"]>>");
        builder.Writeln("<<[model.HtmlA] -html>>");
        builder.Writeln("<</if>>");

        builder.Writeln("<<if [model.Category == \"B\"]>>");
        builder.Writeln("<<[model.HtmlB] -html>>");
        builder.Writeln("<</if>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new()
        {
            Category = "A", // Change to "B" to include the other fragment.
            HtmlA = "<p style=\"color:blue;\">This is HTML fragment <b>A</b>.</p>",
            HtmlB = "<p style=\"color:green;\">This is HTML fragment <b>B</b>.</p>"
        };

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Use the root name "model" because the template tags reference it.
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        doc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Public data model used by the LINQ Reporting engine.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Determines which HTML fragment will be inserted.
    public string Category { get; set; } = string.Empty;

    // HTML fragment for category A.
    public string HtmlA { get; set; } = string.Empty;

    // HTML fragment for category B.
    public string HtmlB { get; set; } = string.Empty;
}
