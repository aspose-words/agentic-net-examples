using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a template document that references a missing enum member.
        DocumentBuilder builder = new DocumentBuilder();
        builder.Writeln("<<[MissingEnum.Value]>>"); // This will be missing in the data source.

        // Configure the ReportingEngine to treat missing members as null literals.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            // An empty string will be printed for missing members (no exception will be thrown).
            MissingMemberMessage = ""
        };

        // Build the report using an empty data source (no enum provided).
        engine.BuildReport(builder.Document, new DataSet());

        // Save the resulting document.
        builder.Document.Save("Output.docx");
    }
}
