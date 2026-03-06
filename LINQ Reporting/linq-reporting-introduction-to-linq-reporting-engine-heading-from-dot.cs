using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOT template that contains a placeholder for the heading.
        // Example placeholder in the template: <<[model.Heading]>>
        Document template = new Document("Template.dot");

        // Create a simple data source with a property that will be inserted into the template.
        var data = new ReportData
        {
            Heading = "LINQ Reporting Introduction to LINQ Reporting Engine"
        };

        // Initialize the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report, binding the data source to the name "model" used in the template.
        engine.BuildReport(template, data, "model");

        // Save the populated document.
        template.Save("Output.docx");
    }

    // Plain .NET class used as the data source for the report.
    public class ReportData
    {
        public string Heading { get; set; }
    }
}
