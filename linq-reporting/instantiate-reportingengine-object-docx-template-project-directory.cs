using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Path to the DOCX template located in the project directory.
        // Adjust the file name as needed.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");

        // Load the template document using the Document(string) constructor.
        Document templateDoc = new Document(templatePath);

        // Instantiate the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();

        // The engine is now ready to build a report with the loaded template.
        // Example usage (optional):
        // bool success = engine.BuildReport(templateDoc, dataSourceObject);
    }
}
