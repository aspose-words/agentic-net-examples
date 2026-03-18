using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load a template document (replace the path with your actual template file)
        Document doc = new Document("Template.docx");

        // Create a ReportingEngine instance
        ReportingEngine engine = new ReportingEngine();

        // Register System.Math so its static members can be accessed from the template
        engine.KnownTypes.Add(typeof(System.Math));

        // Build the report. An empty anonymous object is used here because the example
        // focuses only on the known‑type registration.
        engine.BuildReport(doc, new { });

        // Save the generated document
        doc.Save("Result.docx");
    }
}
