using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and insert a reporting template that references a missing member.
        DocumentBuilder builder = new DocumentBuilder();
        builder.Writeln("<<[missingObject.First().id]>>");
        builder.Writeln("<<foreach [in missingObject]>><<[id]>><</foreach>>");

        // Configure the reporting engine to allow missing members and specify the message to display.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "Missed"
        };

        // Build the report using an empty DataSet (no data for the missing object).
        engine.BuildReport(builder.Document, new DataSet(), string.Empty);

        // Save the resulting document as plain text using TxtSaveOptions.
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        builder.Document.Save("ReportWithMissingMembers.txt", saveOptions);
    }
}
