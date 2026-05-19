using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Define file names in the current working directory.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string reportPath   = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the Sections collection.
        builder.Writeln("<<foreach [sec in Sections]>>");

        // Apply a background color to the whole content of the section.
        // The color is taken from the data source (sec.Color) and can be a name or a hex code.
        builder.Writeln("<<backColor [sec.Color]>>");
        // Section title placeholder.
        builder.Writeln("<<[sec.Title]>>");
        // Close the backColor tag.
        builder.Writeln("<</backColor>>");

        // End the foreach loop.
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
                new SectionInfo { Title = "First Section",  Color = "LightYellow" },
                new SectionInfo { Title = "Second Section", Color = "#FFCCCC" },
                new SectionInfo { Title = "Third Section",  Color = "LightGreen" }
            }
        };

        // -----------------------------------------------------------------
        // 3. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        reportDoc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<SectionInfo> Sections { get; set; } = new();
}

public class SectionInfo
{
    public string Title { get; set; } = string.Empty;
    public string Color { get; set; } = string.Empty;
}
