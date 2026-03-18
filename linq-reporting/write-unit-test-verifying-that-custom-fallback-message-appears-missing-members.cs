using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace Aspose.Words.Reporting.Tests
{
    class Program
    {
        static void Main()
        {
            try
            {
                CustomFallbackMessageIsInsertedForMissingMembers();
                Console.WriteLine("Test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
                Environment.Exit(1);
            }
        }

        // The original NUnit test has been converted to a plain method so the project does not need a reference to the NUnit framework.
        static void CustomFallbackMessageIsInsertedForMissingMembers()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a template expression that references a missing member.
            // The expression will be replaced by the fallback message when the member is not found.
            builder.Writeln("<<[missingObject.First().id]>>");

            // Configure the reporting engine to allow missing members and provide a custom message.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers,
                MissingMemberMessage = "Missed"
            };

            // Build the report using an empty DataSet as the data source.
            engine.BuildReport(doc, new DataSet(), "");

            // Verify that the resulting document contains the custom fallback message.
            string documentText = doc.GetText(); // Retrieves all visible text from the document.
            if (!documentText.Contains("Missed"))
                throw new Exception($"Expected fallback message 'Missed' was not found. Document text: '{documentText}'");
        }
    }
}
