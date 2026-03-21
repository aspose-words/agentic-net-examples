using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class EmptyDataSource { }

class Program
{
    static void Main()
    {
        // Create a simple template document containing a LINQ Reporting Engine expression
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Report generated at: {{DateTime.Now}}");

        // Create a reporting engine instance
        ReportingEngine engine = new ReportingEngine();

        // Add System.DateTime to the set of known types so the template can access static members (e.g., DateTime.Now)
        engine.KnownTypes.Add(typeof(DateTime));

        // Build the report – a visible (public) empty class is used as the data source
        engine.BuildReport(doc, new EmptyDataSource());

        // Save the populated document
        doc.Save("ReportWithDateTime.docx");
    }
}
