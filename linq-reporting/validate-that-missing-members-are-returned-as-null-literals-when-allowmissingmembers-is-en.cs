using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare file paths in the current working directory.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        string resultPath = Path.Combine(Environment.CurrentDirectory, "Result.docx");

        // -------------------------------------------------
        // 1. Create a template document with LINQ Reporting tags
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Tag that tries to access a missing member.
        builder.Writeln("<<[missingObject.First().id]>>");

        // Foreach loop over a missing collection – also missing members.
        builder.Writeln("<<foreach [in missingObject]>><<[id]>><</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template (simulating a real scenario)
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Configure the ReportingEngine to allow missing members
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            // Optional: customize the message shown for a missing plain member.
            // Leaving it empty makes the engine output a null literal (empty string).
            MissingMemberMessage = string.Empty
        };

        // Use an empty DataSet as the data source – it contains no members.
        DataSet emptyData = new DataSet();

        // Build the report. The third parameter is the data source name; an empty string means we don't reference the object itself.
        engine.BuildReport(doc, emptyData, "");

        // -------------------------------------------------
        // 4. Save the generated report
        // -------------------------------------------------
        doc.Save(resultPath);

        // -------------------------------------------------
        // 5. Verify the output – missing members should be empty.
        // -------------------------------------------------
        string resultText = doc.GetText();

        Console.WriteLine("=== Generated Report Text ===");
        Console.WriteLine(resultText);
        Console.WriteLine("=== Validation ===");
        if (string.IsNullOrWhiteSpace(resultText))
        {
            Console.WriteLine("Success: Missing members were rendered as null literals (empty).");
        }
        else
        {
            Console.WriteLine("Failure: Unexpected content found.");
        }
    }
}
