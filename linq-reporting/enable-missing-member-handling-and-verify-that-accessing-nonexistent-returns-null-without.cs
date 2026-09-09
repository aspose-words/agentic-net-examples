using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class MissingMemberExample
{
    public static void Main()
    {
        // Create a template document with a tag that references a missing member.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<[nonexistent]>>"); // This member does not exist in the data source.
        templateDoc.Save(templatePath);

        // Load the template back from disk.
        Document doc = new Document(templatePath);

        // Configure the reporting engine to treat missing members as null.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };
        // Optional: customize the message printed for missing members (not required for null handling).
        engine.MissingMemberMessage = "null";

        // Build the report using an empty data source (object with no members).
        bool success = engine.BuildReport(doc, new object(), "data");

        // Verify that the missing member was treated as null (empty string in the output).
        string resultText = doc.GetText().Trim();
        Console.WriteLine($"Build succeeded: {success}");
        Console.WriteLine($"Resulting document text: '{resultText}'");
        // Expected output: an empty string between the quotes.
    }
}
