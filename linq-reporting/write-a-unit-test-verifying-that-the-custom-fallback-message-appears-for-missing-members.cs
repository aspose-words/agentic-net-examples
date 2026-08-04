using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that references a missing member.
        // The engine will replace this with the custom fallback message.
        builder.Writeln("<<[missingObject.First().id]>>");

        // Configure the ReportingEngine to allow missing members and set the fallback message.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = "Missed";

        // Build the report. The data source is irrelevant here; we pass an empty DataSet.
        bool success = engine.BuildReport(doc, new DataSet(), "");

        // Verify that the report was built successfully and that the fallback message appears.
        string resultText = doc.GetText();

        bool containsFallback = resultText.Contains("Missed");

        if (success && containsFallback)
        {
            Console.WriteLine("Test passed: custom fallback message was inserted.");
        }
        else
        {
            Console.WriteLine("Test failed: fallback message not found.");
        }

        // Save the generated document for manual inspection (optional).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MissingMemberFallback.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
