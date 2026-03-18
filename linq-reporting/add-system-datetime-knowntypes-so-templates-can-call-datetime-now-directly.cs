using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load a template document that contains a LINQ Reporting Engine expression like {{DateTime.Now}}
        Document doc = new Document("Template.docx");

        // Create a reporting engine instance
        ReportingEngine engine = new ReportingEngine();

        // Add System.DateTime to the set of known types so the template can access static members (e.g., DateTime::Now)
        engine.KnownTypes.Add(typeof(DateTime));

        // Build the report – in this simple example we do not need a data source, so we pass null
        engine.BuildReport(doc, null);

        // Save the populated document
        doc.Save("ReportWithDateTime.docx");
    }
}
