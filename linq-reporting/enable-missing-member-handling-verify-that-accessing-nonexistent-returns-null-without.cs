using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new document containing a template expression that refers to a non‑existent member.
        DocumentBuilder builder = new DocumentBuilder();
        builder.Writeln("<<[nonexistent]>>"); // <<[nonexistent]>> will be processed by the ReportingEngine.

        // Configure the ReportingEngine to treat missing members as null literals.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            // MissingMemberMessage defaults to an empty string; setting it explicitly makes the intent clear.
            MissingMemberMessage = ""
        };

        // Build the report using an empty data source (no objects are needed for this test).
        engine.BuildReport(builder.Document, new object());

        // After building, the missing member should be rendered as an empty string (null literal).
        string output = builder.Document.GetText().Trim();

        // Verify that the output is empty; otherwise the missing‑member handling failed.
        if (output != string.Empty)
            throw new InvalidOperationException("Missing member was not treated as null.");

        // Save the resulting document (optional, useful for manual inspection).
        builder.Document.Save("MissingMemberResult.docx");
    }
}
