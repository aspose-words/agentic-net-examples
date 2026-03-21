using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class EmptyDataSource
{
    // No members needed; this class is only used to satisfy the data source requirement.
}

class Program
{
    static void Main()
    {
        // Create an empty document (or load a template if you have one)
        Document doc = new Document();

        // Create a ReportingEngine instance
        ReportingEngine engine = new ReportingEngine();

        // Register System.Math so its static members can be accessed from the template
        engine.KnownTypes.Add(typeof(System.Math));

        // Build the report using a visible data source type
        engine.BuildReport(doc, new EmptyDataSource());

        // Save the generated document
        doc.Save("Result.docx");
    }
}
