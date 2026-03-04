using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new blank document and add a template that references a missing object.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("<<[missingObject.First().id]>>");
        builder.Writeln("<<foreach [in missingObject]>><<[id]>><</foreach>>");

        // Configure the ReportingEngine:
        // - Allow missing members so the engine does not throw.
        // - Provide a custom message that will be printed for each missing member.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "Missed"
        };

        // Build the report using an empty data source (no actual data is needed for this demo).
        engine.BuildReport(doc, new DataSet(), "");

        // Save the generated document.
        doc.Save("ReportWithMissingMembers.docx");
    }
}
