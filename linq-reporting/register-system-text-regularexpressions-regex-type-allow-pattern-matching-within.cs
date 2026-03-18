using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the template document that contains reporting engine tags.
        Document template = new Document("Template.docx");

        // Create a reporting engine instance.
        ReportingEngine engine = new ReportingEngine();

        // Register the Regex type so that its static members can be used inside template expressions.
        engine.KnownTypes.Add(typeof(Regex));

        // Build the report using a data source (replace with your actual data object).
        object dataSource = GetDataSource();
        engine.BuildReport(template, dataSource);

        // Save the generated report.
        template.Save("ReportOutput.docx");
    }

    // Placeholder method to provide a data source for the report.
    static object GetDataSource()
    {
        // Return any object that matches the fields used in the template.
        // For demonstration, an anonymous object with a simple property is returned.
        return new { Name = "John Doe", Age = 30 };
    }
}
