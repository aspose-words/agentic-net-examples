using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the example files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "LinQReportingExample");
        Directory.CreateDirectory(workDir);

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "report.docx");

        // -----------------------------------------------------------------
        // 1. Build the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Insert a heading for each section using a foreach tag.
        builder.Writeln("<<foreach [sec in Sections]>>");
        // The Letter property converts the integer Number to an uppercase letter.
        builder.Writeln("<<[sec.Letter]>>. <<[sec.Title]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new()
        {
            Sections = new()
            {
                new Section { Number = 1, Title = "Introduction" },
                new Section { Number = 2, Title = "Methodology" },
                new Section { Number = 3, Title = "Results" },
                new Section { Number = 4, Title = "Conclusion" }
            }
        };

        // -----------------------------------------------------------------
        // 3. Load the template and generate the report.
        // -----------------------------------------------------------------
        Document doc = new(templatePath);
        ReportingEngine engine = new();
        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Collection of sections to be reported.
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    // Numeric identifier of the section.
    public int Number { get; set; }

    // Title of the section.
    public string Title { get; set; } = string.Empty;

    // Converts the numeric section number to an uppercase alphabetic letter (A, B, C, ...).
    public string Letter => ((char)('A' + Number - 1)).ToString();
}
