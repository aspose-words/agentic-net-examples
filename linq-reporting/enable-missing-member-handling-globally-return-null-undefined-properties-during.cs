using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a template document that references a member which does not exist in the data source.
        DocumentBuilder builder = new DocumentBuilder();
        builder.Writeln("<<[missingObject.First().Id]>>");
        builder.Writeln("<<foreach [in missingObject]>><<[Id]>><</foreach>>");

        // Configure the ReportingEngine to treat missing members as null literals.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            // An empty string means the engine will output nothing for a missing member (null literal).
            MissingMemberMessage = ""
        };

        // Build the report using an empty data source; missing members will be rendered as null.
        DataSet emptyData = new DataSet();
        engine.BuildReport(builder.Document, emptyData, "");

        // Save the resulting document.
        builder.Document.Save("MissingMembersHandled.docx");
    }
}
