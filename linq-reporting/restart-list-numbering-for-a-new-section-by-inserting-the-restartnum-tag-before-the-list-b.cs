using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        ReportModel model = new()
        {
            Sections = new()
            {
                new Section
                {
                    Title = "Section A",
                    Items = new() { "Item A1", "Item A2", "Item A3" }
                },
                new Section
                {
                    Title = "Section B",
                    Items = new() { "Item B1", "Item B2" }
                }
            }
        };

        // Create the template document programmatically
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin outer foreach over sections
        builder.Writeln("<<foreach [section in Model.Sections]>>");
        // Section title
        builder.Writeln("<<[section.Title]>>");
        // Numbered list of items with restartNum before inner foreach
        builder.Writeln("1. <<restartNum>><<foreach [item in section.Items]>> <<[item]>> <</foreach>>");
        // End outer foreach
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "Model");

        // Save the generated report
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Data model classes
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    public string Title { get; set; } = string.Empty;
    public List<string> Items { get; set; } = new();
}
