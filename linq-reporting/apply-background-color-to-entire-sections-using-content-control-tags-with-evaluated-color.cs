using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the Sections collection.
        builder.Writeln("<<foreach [section in Sections]>>");

        // Apply a background color to the whole block using the backColor tag.
        // The color is taken from the data source (section.Color).
        builder.Writeln("<<backColor [section.Color]>>");

        // Content that will be colored – the section title.
        builder.Writeln("<<[section.Title]>>");

        // Close the backColor tag.
        builder.Writeln("<</backColor>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Data model with two sections, each having a title and a background color.
        var model = new ReportModel
        {
            Sections = new List<SectionInfo>
            {
                new SectionInfo { Title = "First Section", Color = "LightYellow" },
                new SectionInfo { Title = "Second Section", Color = "LightBlue" }
            }
        };

        // -----------------------------------------------------------------
        // 3. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options

        // The root object name in the template is "model".
        engine.BuildReport(loadedTemplate, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        loadedTemplate.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes used by the LINQ Reporting engine.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Collection of sections to be iterated over in the template.
    public List<SectionInfo> Sections { get; set; } = new();
}

public class SectionInfo
{
    // Title displayed for the section.
    public string Title { get; set; } = string.Empty;

    // Background color name or HTML color code (e.g., "LightYellow" or "#FFCC00").
    public string Color { get; set; } = string.Empty;
}
