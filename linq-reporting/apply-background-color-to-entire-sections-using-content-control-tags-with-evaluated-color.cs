using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare the template document with LINQ Reporting tags.
        var templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin foreach over Sections collection.
        builder.Writeln("<<foreach [sec in Sections]>>");
        // Apply background color to the whole section using backColor tag.
        builder.Writeln("<<backColor [sec.BgColor]>>");
        // Section title.
        builder.Writeln("<<[sec.Title]>>");
        // Section body.
        builder.Writeln("<<[sec.Body]>>");
        // Close backColor tag.
        builder.Writeln("<</backColor>>");
        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Create sample data model.
        var model = new ReportModel
        {
            Sections = new()
            {
                new Section
                {
                    Title = "Section 1",
                    Body = "Content of section 1.",
                    BgColor = "LightYellow"
                },
                new Section
                {
                    Title = "Section 2",
                    Body = "Content of section 2.",
                    BgColor = "#FFCCCC"
                },
                new Section
                {
                    Title = "Section 3",
                    Body = "Content of section 3.",
                    BgColor = "LightGreen"
                }
            }
        };

        // Build the report using LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var outputPath = "output.docx";
        doc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    public string Title { get; set; } = string.Empty;
    public string Body { get; set; } = string.Empty;
    public string BgColor { get; set; } = string.Empty;
}
