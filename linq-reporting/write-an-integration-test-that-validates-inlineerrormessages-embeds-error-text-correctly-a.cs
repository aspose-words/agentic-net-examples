using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorMessagesDemo
{
    // Simple data model used as the root object for the report.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
    }

    class Program
    {
        static void Main()
        {
            // Prepare file paths.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "DemoFiles");
            Directory.CreateDirectory(workDir);
            string templatePath = Path.Combine(workDir, "template.docx");
            string resultPath = Path.Combine(workDir, "result.docx");

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // The template contains a valid tag <<[person.Name]>> and an invalid tag
            // <<[person.MissingProperty]>> which will trigger an inline error message.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Hello <<[person.Name]>>!");
            builder.Writeln("This line contains a missing member: <<[person.MissingProperty]>>");
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (simulating a real-world scenario where the
            //    template is stored on disk).
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            Person model = new Person { Name = "John Doe" };

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine to use InlineErrorMessages.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report. The method returns a bool indicating whether the
            // template was parsed without errors (when InlineErrorMessages is set).
            bool success = engine.BuildReport(reportDoc, model, "person");

            // Save the generated document.
            reportDoc.Save(resultPath);

            // -----------------------------------------------------------------
            // 5. Validate the outcome.
            //    - The success flag should be false because the template contained
            //      an invalid expression.
            //    - The resulting document should contain an inline error message.
            // -----------------------------------------------------------------
            string resultText = reportDoc.GetText();

            bool containsErrorMessage = resultText.Contains("Error", StringComparison.OrdinalIgnoreCase);

            Console.WriteLine($"BuildReport success flag: {success}");
            Console.WriteLine($"Document contains inline error message: {containsErrorMessage}");

            // Simple assertions to emulate an integration test.
            if (success)
                Console.WriteLine("FAIL: Expected success flag to be false due to missing member.");
            else
                Console.WriteLine("PASS: Success flag correctly indicates parsing errors.");

            if (containsErrorMessage)
                Console.WriteLine("PASS: Inline error message was embedded in the document.");
            else
                Console.WriteLine("FAIL: Inline error message was not found in the document.");

            // Clean up (optional).
            // File.Delete(templatePath);
            // File.Delete(resultPath);
        }
    }
}
