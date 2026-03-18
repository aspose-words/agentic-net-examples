using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a template document that contains a syntax error.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // This field will cause a syntax error because the object does not exist.
        builder.Writeln("<<[missingObject.First().Id]>>");

        // Configure the ReportingEngine to inline error messages.
        ReportingEngine engine = new ReportingEngine
        {
            // Enable inline error messages and allow missing members.
            Options = ReportBuildOptions.InlineErrorMessages | ReportBuildOptions.AllowMissingMembers,
            // The placeholder that will be inserted where an error occurs.
            MissingMemberMessage = "<<error>>"
        };

        // Build the report using an empty data source (no data needed for this example).
        engine.BuildReport(doc, new DataSet(), "");

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
