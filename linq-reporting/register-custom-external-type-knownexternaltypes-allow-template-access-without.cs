using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load a template document that contains LINQ Reporting Engine expressions.
        // The file "Template.docx" should exist in the application directory.
        Document template = new Document("Template.docx");

        // Create an instance of the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();

        // Register a custom external type so that its static members can be accessed
        // directly from the template without using reflection.
        // Example: expose the System.Math class.
        engine.KnownTypes.Add(typeof(Math));

        // Build the report. In this example we pass an empty anonymous object as the data source
        // because the template may only use static members of the registered type.
        engine.BuildReport(template, new { });

        // Save the populated document.
        template.Save("Result.docx");
    }
}
