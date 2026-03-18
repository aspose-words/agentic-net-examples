using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class InlineErrorDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a template expression that will cause a syntax error.
        // The expression references a non‑existent member "NonExistent".
        builder.Writeln("<<[data.NonExistent]>>");

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            // Enable inline error messages and allow missing members.
            Options = ReportBuildOptions.InlineErrorMessages | ReportBuildOptions.AllowMissingMembers,
            // The placeholder that will be inserted where the error occurs.
            MissingMemberMessage = "<<error>>"
        };

        // Build the report using an empty data source (no "data" object is provided).
        // The engine will inline the "<<error>>" placeholder into the output.
        engine.BuildReport(doc, new object());

        // Save the resulting document.
        doc.Save("InlineErrorResult.docx");
    }
}
