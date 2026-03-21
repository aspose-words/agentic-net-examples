using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class BatchReportProcessor
{
    public static void Main()
    {
        // Use directories relative to the current working directory
        string baseDir = Directory.GetCurrentDirectory();
        string templatesDir = Path.Combine(baseDir, "Templates");
        string outputDir = Path.Combine(baseDir, "Reports");

        // Ensure directories exist
        Directory.CreateDirectory(templatesDir);
        Directory.CreateDirectory(outputDir);

        // If there are no templates, create a simple one for demonstration
        if (Directory.GetFiles(templatesDir, "*.docx").Length == 0)
        {
            string sampleTemplatePath = Path.Combine(templatesDir, "SampleTemplate.docx");
            var sampleDoc = new Document();
            var builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("Report Title: <<Title>>");
            builder.Writeln("Generated on: <<GeneratedOn>>");
            sampleDoc.Save(sampleTemplatePath);
        }

        // Configure the ReportingEngine with identical options for all reports
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs | ReportBuildOptions.AllowMissingMembers
        };

        // Example data source that will be used for every template
        var data = new ReportData
        {
            Title = "Monthly Summary",
            GeneratedOn = DateTime.Now
        };

        // Process each .docx template in the source folder
        foreach (string templatePath in Directory.GetFiles(templatesDir, "*.docx"))
        {
            // Load the template document
            var doc = new Document(templatePath);

            // Populate the template with the data source
            engine.BuildReport(doc, data);

            // Construct the output file name
            string outputFileName = Path.GetFileNameWithoutExtension(templatePath) + "_Report.docx";
            string outputPath = Path.Combine(outputDir, outputFileName);

            // Save the populated document
            doc.Save(outputPath);
        }

        Console.WriteLine($"Processing complete. Reports saved to: {outputDir}");
    }
}

// Simple POCO class used as the data source for the reporting engine
public class ReportData
{
    public string Title { get; set; }
    public DateTime GeneratedOn { get; set; }
}
