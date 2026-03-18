using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class BatchReportProcessor
{
    static void Main()
    {
        // Directory containing the DOCX templates
        string templatesDir = @"C:\Templates";

        // Directory where the generated reports will be saved
        string outputDir = @"C:\Reports";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputDir);

        // Create a ReportingEngine instance and configure identical options for all reports
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs | ReportBuildOptions.AllowMissingMembers;

        // Example data source that will be used for every template
        ReportData data = new ReportData
        {
            Title = "Monthly Summary",
            GeneratedOn = DateTime.Now
        };

        // Loop through each .docx template in the source folder
        foreach (string templatePath in Directory.GetFiles(templatesDir, "*.docx"))
        {
            // Load the template document (lifecycle rule: use Document(string) constructor)
            Document doc = new Document(templatePath);

            // Populate the template with the data source using the configured ReportingEngine
            engine.BuildReport(doc, data);

            // Construct the output file name
            string outputFileName = Path.GetFileNameWithoutExtension(templatePath) + "_Report.docx";
            string outputPath = Path.Combine(outputDir, outputFileName);

            // Save the populated document (lifecycle rule: use Document.Save(string) method)
            doc.Save(outputPath);
        }
    }

    // Simple POCO class used as the data source for the reporting engine
    public class ReportData
    {
        public string Title { get; set; }
        public DateTime GeneratedOn { get; set; }
    }
}
