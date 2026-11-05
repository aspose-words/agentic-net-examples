using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string jsonPath = "sampleData.json";
        var sampleData = new ReportRoot
        {
            Sections = new List<ReportSection>
            {
                new()
                {
                    Title = "Generated Section 1",
                    Paragraphs = new List<string> { "First paragraph.", "", "Third paragraph." },
                    Items = new List<ReportItem>
                    {
                        new() { Name = "Item A" },
                        new() { Name = "" },
                        new() { Name = "Item C" }
                    }
                },
                new()
                {
                    Title = "Generated Section 2",
                    Paragraphs = new List<string> { "", "Only paragraph with content." },
                    Items = new List<ReportItem>()
                }
            }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Create JSON-driven template.
        string jsonTemplatePath = "jsonTemplate.docx";
        var jsonTemplate = new Document();
        var jsonBuilder = new DocumentBuilder(jsonTemplate);
        jsonBuilder.Writeln("<<foreach [sec in Sections]>>");
        jsonBuilder.Writeln("<<[sec.Title]>>");
        jsonBuilder.Writeln("<<foreach [para in sec.Paragraphs]>>");
        jsonBuilder.Writeln("<<[para]>>");
        jsonBuilder.Writeln("<</foreach>>");
        jsonBuilder.Writeln("<<foreach [it in sec.Items]>>");
        jsonBuilder.Writeln("Item: <<[it.Name]>>");
        jsonBuilder.Writeln("<</foreach>>");
        jsonBuilder.Writeln("<</foreach>>");
        jsonTemplate.Save(jsonTemplatePath);

        // Build report for JSON sections with removal of empty paragraphs.
        var jsonDoc = new Document(jsonTemplatePath);
        var jsonEngine = new ReportingEngine();
        jsonEngine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        jsonEngine.BuildReport(jsonDoc, sampleData, "root");
        string jsonReportPath = "jsonReport.docx";
        jsonDoc.Save(jsonReportPath);

        // Create static content template.
        string staticTemplatePath = "staticTemplate.docx";
        var staticDoc = new Document();
        var staticBuilder = new DocumentBuilder(staticDoc);
        staticBuilder.Writeln("=== Static Content Section ===");
        staticBuilder.Writeln("This paragraph should remain even if empty.");
        staticBuilder.Writeln(""); // Intentionally empty paragraph that should stay.
        staticBuilder.Writeln("End of static content.");
        staticDoc.Save(staticTemplatePath);

        // Load static document and append JSON report sections.
        var finalDoc = new Document(staticTemplatePath);
        var jsonReportDoc = new Document(jsonReportPath);
        finalDoc.AppendDocument(jsonReportDoc, ImportFormatMode.KeepSourceFormatting);

        // Save final combined document.
        string outputPath = "finalReport.docx";
        finalDoc.Save(outputPath);
    }
}

// Data model classes.
public class ReportRoot
{
    public List<ReportSection> Sections { get; set; } = new();
}

public class ReportSection
{
    public string Title { get; set; } = string.Empty;
    public List<string> Paragraphs { get; set; } = new();
    public List<ReportItem> Items { get; set; } = new();
}

public class ReportItem
{
    public string Name { get; set; } = string.Empty;
}
