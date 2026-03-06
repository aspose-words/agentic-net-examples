using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a mail‑merge style template that references a member which does not exist.
        // The ReportingEngine will handle the missing member according to the options we set later.
        builder.Writeln("<<[missingObject.First().Id]>>");
        builder.Writeln("<<foreach [in missingObject]>><<[Id]>><</foreach>>");

        // Configure the ReportingEngine:
        // - AllowMissingMembers lets the engine continue when a member is not found.
        // - MissingMemberMessage defines the text that will be printed in place of the missing member.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "Missing"
        };

        // Build the report using an empty data source (no actual data is required for this demo).
        engine.BuildReport(doc, new DataSet(), "");

        // Save the resulting document in DOCM format (Word macro‑enabled document).
        doc.Save("Report.docm", SaveFormat.Docm);
    }
}
