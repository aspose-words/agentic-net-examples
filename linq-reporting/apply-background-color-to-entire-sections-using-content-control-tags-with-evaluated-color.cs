using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Step 1: Create the template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the Sections collection.
        builder.Writeln("<<foreach [section in Sections]>>");

        // Apply background color using a backColor tag whose expression is evaluated per section.
        builder.Writeln("<<backColor [section.Color]>>");
        // Section title.
        builder.Writeln("<<[section.Title]>>");
        // Sample body text for the section.
        builder.Writeln("This is the body of the <<[section.Title]>> section.");
        // Close the backColor tag.
        builder.Writeln("<</backColor>>");

        // Insert a section break so each iteration starts on a new page.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Step 2: Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Step 3: Prepare the data model.
        ReportModel model = new ReportModel
        {
            Sections = new List<SectionInfo>
            {
                new SectionInfo { Title = "Introduction", Color = "LightYellow" },
                new SectionInfo { Title = "Details", Color = "#FFCCCB" }, // Light red.
                new SectionInfo { Title = "Conclusion", Color = "LightGray" }
            }
        };

        // Step 4: Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // No special options required.
        engine.BuildReport(reportDoc, model, "model");

        // Step 5: Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<SectionInfo> Sections { get; set; } = new();
}

public class SectionInfo
{
    public string Title { get; set; } = string.Empty;
    public string Color { get; set; } = string.Empty; // Can be a known color name or HTML hex code.
}
