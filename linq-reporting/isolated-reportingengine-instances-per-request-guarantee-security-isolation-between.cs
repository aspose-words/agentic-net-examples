using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class ReportProcessor
{
    // Processes a single report request.
    // Each call creates its own ReportingEngine instance, ensuring isolation between users.
    public void ProcessReport(string templatePath, object dataSource, string outputPath)
    {
        // Load the template document from file.
        Document template = new Document(templatePath);

        // Create a fresh ReportingEngine for this request.
        ReportingEngine engine = new ReportingEngine();

        // Build the report by populating the template with the supplied data source.
        engine.BuildReport(template, dataSource);

        // Create appropriate save options for the desired output format.
        SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Docx);

        // Save the generated report to the specified path.
        template.Save(outputPath, saveOptions);
    }
}

public static class Program
{
    // Entry point required for a console application.
    public static void Main(string[] args)
    {
        // Example usage – replace with real paths and data source as needed.
        string templatePath = "template.docx"; // path to the Aspose.Words template
        string outputPath = "report.docx";     // where the generated report will be saved
        var dataSource = new { Name = "John Doe", Date = DateTime.Today }; // simple anonymous object as data source

        var processor = new ReportProcessor();
        processor.ProcessReport(templatePath, dataSource, outputPath);

        Console.WriteLine($"Report generated at: {outputPath}");
    }
}
