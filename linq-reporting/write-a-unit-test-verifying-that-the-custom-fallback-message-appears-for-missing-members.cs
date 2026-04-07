using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Entry point of the console application.
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that references a missing member.
        // The engine will replace this with the custom fallback message.
        builder.Writeln("<<[missingObject.Name]>>");

        // Save the template (optional, shown for completeness).
        const string templatePath = "template.docx";
        doc.Save(templatePath);

        // Load the template back (demonstrates the load step).
        Document template = new Document(templatePath);

        // Configure the ReportingEngine.
        ReportingEngine engine = new ReportingEngine
        {
            // Treat missing members as null literals.
            Options = ReportBuildOptions.AllowMissingMembers,
            // Custom message to display when a member is missing.
            MissingMemberMessage = "Custom fallback"
        };

        // Build the report using an empty DataSet as the data source.
        // The third parameter (data source name) is empty because we do not reference the object itself.
        bool success = engine.BuildReport(template, new DataSet(), "");

        // Retrieve the resulting text from the document.
        string resultText = template.GetText();

        // Verify that the custom fallback message appears in the output.
        bool testPassed = resultText.Contains("Custom fallback");

        // Output the result and test status.
        Console.WriteLine("Report generation success flag: " + success);
        Console.WriteLine("Resulting document text:");
        Console.WriteLine(resultText);
        Console.WriteLine("Unit test " + (testPassed ? "PASSED" : "FAILED"));
    }
}
