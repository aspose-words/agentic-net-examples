using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and insert a template that references a missing object.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("<<[MissingObject.First().Id]>>");
        builder.Writeln("<<foreach [in MissingObject]>><<[Id]>><</foreach>>");

        // Configure the reporting engine to allow missing members and set a custom message.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };
        engine.MissingMemberMessage = "Missed";

        // Build the report using an empty data source (no MissingObject present).
        engine.BuildReport(doc, new DataSet(), string.Empty);

        // Save the resulting document as plain text.
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        doc.Save("Report.txt", saveOptions);
    }
}
