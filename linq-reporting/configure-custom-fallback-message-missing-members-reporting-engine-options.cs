using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a template document that references a member which does not exist.
        DocumentBuilder builder = new DocumentBuilder();
        builder.Writeln("<<[missingObject.First().id]>>");
        builder.Writeln("<<foreach [in missingObject]>><<[id]>><</foreach>>");

        // Configure the reporting engine:
        // - Allow missing members so the engine does not throw an exception.
        // - Provide a custom fallback message that will be printed instead of the missing member.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "Member not found"
        };

        // Build the report using an empty data source (no objects named 'missingObject').
        engine.BuildReport(builder.Document, new DataSet());

        // Save the generated document.
        builder.Document.Save("MissingMemberReport.docx");
    }
}
