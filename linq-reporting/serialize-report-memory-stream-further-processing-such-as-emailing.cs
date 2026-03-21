using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class ReportData
{
    public string Title { get; set; }
    public string Date { get; set; }
    public double Total { get; set; }
}

class ReportToMemoryStream
{
    // Generates a report from a template and returns it as a MemoryStream.
    public static MemoryStream GenerateReport(string templatePath, ReportData dataSource)
    {
        // Load the template document from file.
        Document doc = new Document(templatePath);

        // Populate the template with data using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource);

        // Save the populated document to a memory stream in DOCX format.
        MemoryStream stream = new MemoryStream();
        doc.Save(stream, SaveFormat.Docx);
        stream.Position = 0; // Rewind for reading.

        return stream;
    }

    static void Main()
    {
        // Create a temporary template file with simple placeholders.
        string tempDir = Path.GetTempPath();
        string templatePath = Path.Combine(tempDir, "ReportTemplate.docx");

        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Title: {{Title}}");
        builder.Writeln("Date: {{Date}}");
        builder.Writeln("Total: {{Total}}");
        templateDoc.Save(templatePath);
        // DocumentBuilder does not implement IDisposable; no need to call Dispose.

        // Example data source.
        var data = new ReportData
        {
            Title = "Quarterly Sales",
            Date = DateTime.Now.ToString("d"),
            Total = 12345.67
        };

        // Generate the report into a memory stream.
        using MemoryStream reportStream = GenerateReport(templatePath, data);

        // Write the stream contents to a temporary output file.
        string outputPath = Path.Combine(tempDir, "Report.docx");
        using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        {
            reportStream.CopyTo(file);
        }

        Console.WriteLine($"Report generated at: {outputPath}");
    }
}
