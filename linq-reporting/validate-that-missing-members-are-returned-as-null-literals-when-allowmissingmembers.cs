using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class AllowMissingMembersDemo
{
    static void Main()
    {
        // Create a new document and insert a template expression that references a missing member.
        DocumentBuilder builder = new DocumentBuilder();
        // The expression <<[missingObject.id]>> will try to access the "id" property of a non‑existent object.
        builder.Writeln("<<[missingObject.id]>>");

        // Configure the reporting engine to treat missing members as null literals.
        ReportingEngine engine = new ReportingEngine
        {
            // Enable the option that converts missing members to null.
            Options = ReportBuildOptions.AllowMissingMembers,
            // Optional: specify a custom string to be printed for missing members.
            MissingMemberMessage = "null"
        };

        // Build the report using an empty data source (no "missingObject" defined).
        // An empty DataSet is sufficient because the template does not reference any real data.
        engine.BuildReport(builder.Document, new DataSet(), "");

        // The resulting document should contain the literal defined in MissingMemberMessage.
        string result = builder.Document.GetText().Trim();

        // Simple validation – the output must be exactly "null".
        if (result == "null")
        {
            Console.WriteLine("Success: missing members were rendered as null literals.");
        }
        else
        {
            Console.WriteLine($"Failure: unexpected output \"{result}\".");
        }
    }
}
