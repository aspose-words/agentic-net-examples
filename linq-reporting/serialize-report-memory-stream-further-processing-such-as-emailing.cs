using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class ReportToMemoryStream
{
    // Generates a report from a template and returns it as a MemoryStream.
    public static MemoryStream GenerateReport(string templatePath, object dataSource)
    {
        // Load the template document from file.
        Document doc = new Document(templatePath);

        // Populate the template with data using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource);

        // Create a memory stream that will receive the document bytes.
        MemoryStream stream = new MemoryStream();

        // Save the populated document to the stream in DOCX format.
        // This uses the documented overload: Save(Stream, SaveFormat).
        doc.Save(stream, SaveFormat.Docx);

        // Rewind the stream so that callers can read from the beginning.
        stream.Position = 0;

        return stream;
    }

    // Example entry point demonstrating how to use the method.
    static void Main()
    {
        // Path to the .docx template file.
        string templatePath = @"C:\Templates\ReportTemplate.docx";

        // Example data source – any supported type (anonymous object, DataSet, etc.).
        var data = new
        {
            Title = "Quarterly Sales",
            Date = DateTime.Now,
            Total = 12345.67
        };

        // Generate the report into a memory stream.
        MemoryStream reportStream = GenerateReport(templatePath, data);

        // The stream can now be attached to an email, sent over a network, etc.
        // For illustration, write the stream contents to a physical file.
        using (FileStream file = new FileStream(@"C:\Output\Report.docx", FileMode.Create, FileAccess.Write))
        {
            reportStream.CopyTo(file);
        }

        // Clean up the memory stream when finished.
        reportStream.Dispose();
    }
}
