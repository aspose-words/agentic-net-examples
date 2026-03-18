using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new document and add template tags that reference a missing object.
        DocumentBuilder builder = new DocumentBuilder();
        builder.Writeln("<<[missingObject.First().id]>>");
        builder.Writeln("<<foreach [in missingObject]>><<[id]>><</foreach>>");

        // Initialize the reporting engine and enable the AllowMissingMembers option.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            // Optional: specify a placeholder text for missing members.
            MissingMemberMessage = "Missed"
        };

        // Build the report using an empty DataSet as the data source.
        // Missing members will be treated as null (or the MissingMemberMessage text).
        engine.BuildReport(builder.Document, new DataSet(), "");

        // Save the resulting document.
        builder.Document.Save("ReportWithMissingMembers.docx");
    }
}
