using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = "template.docx";
        string reportPath = "report.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a LINQ Reporting tag that references a non‑existent member.
        // The engine will treat this as a missing member.
        builder.Writeln("<<[nonexistent]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back from the file system.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Configure the ReportingEngine to allow missing members.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            // Treat missing members as null literals.
            Options = ReportBuildOptions.AllowMissingMembers,
            // Optional: customize the message printed for a missing member.
            // Leaving it empty results in an empty string in the output.
            MissingMemberMessage = string.Empty
        };

        // Build the report using an empty data source (any object works).
        // The root name is left empty because the template does not reference it.
        engine.BuildReport(loadedDoc, new object(), "");

        // -----------------------------------------------------------------
        // 4. Verify that the missing member was rendered as an empty string.
        // -----------------------------------------------------------------
        string resultText = loadedDoc.GetText().Trim();
        // The result should be empty (null handling succeeded).
        bool isEmpty = string.IsNullOrEmpty(resultText);
        Console.WriteLine($"Missing member rendered as empty: {isEmpty}");

        // Save the final report.
        loadedDoc.Save(reportPath);
    }
}
