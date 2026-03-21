using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create an empty document (or load a template if you have one).
        Document template = new Document();

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Register a custom external type (e.g., System.Math) so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(Math));

        // Build the report using a visible data source (System.Object instance).
        engine.BuildReport(template, new object());

        // Save the result.
        template.Save("Result.docx");
    }
}
