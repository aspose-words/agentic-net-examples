using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // 1. Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Use the default LINQ Reporting delimiters << and >>.
        // The tag references the root object name "model" and its Greeting property.
        builder.Writeln("<<[model.Greeting]>>");

        // 2. Prepare the data model that will be bound to the template.
        ReportModel model = new ReportModel
        {
            Greeting = "Hello, Aspose.Words LINQ Reporting!"
        };

        // 3. Configure the ReportingEngine (no custom delimiters needed).
        ReportingEngine engine = new ReportingEngine();

        // 4. Build the report. The root object name must match the name used in the template tags.
        engine.BuildReport(template, model, "model");

        // 5. Save the generated report.
        template.Save("CustomDelimiterReport.docx");
    }

    // Simple public data model with a single property.
    public class ReportModel
    {
        public string Greeting { get; set; } = string.Empty;
    }
}
