using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load a DOTX template document.
        // The template can contain expressions like <<[MissingObject.Name]>>.
        Document template = new Document("Template.dotx");

        // Create a ReportingEngine instance.
        ReportingEngine engine = new ReportingEngine
        {
            // Allow the engine to handle missing members without throwing an exception.
            Options = ReportBuildOptions.AllowMissingMembers,

            // Text that will be printed in place of any missing member reference.
            MissingMemberMessage = "Missing"
        };

        // Build the report using an empty data source (no actual data needed for this demo).
        // The third parameter is the name used to reference the data source in the template.
        engine.BuildReport(template, new DataSet(), "");

        // Save the populated document.
        template.Save("Report.docx");
    }
}
