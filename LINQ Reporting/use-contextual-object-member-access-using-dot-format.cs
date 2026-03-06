using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class ReportingEngineExample
{
    static void Main()
    {
        // Load a template document that contains reporting tags.
        Document template = new Document("Template.docx");

        // Create a ReportingEngine instance.
        ReportingEngine engine = new ReportingEngine();

        // Allow missing members in the data source and define a custom message.
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = "Missing";

        // Prepare a data source – an empty DataSet for this example.
        DataSet data = new DataSet();

        // Build the report using the template and the data source.
        engine.BuildReport(template, data, "");

        // Save the generated report.
        template.Save("Result.docx");
    }
}
