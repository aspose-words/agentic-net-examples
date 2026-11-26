using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class InlineErrorMessagesTest
{
    // Simple data model with only one valid property.
    public class Model
    {
        public string ExistingProperty { get; set; } = "Hello World";
        // Note: No MissingProperty defined – this will trigger an inline error.
    }

    public static void Main()
    {
        // Prepare file paths in the current working directory.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a line with a valid expression.
        builder.Writeln("Existing: <<[model.ExistingProperty]>>");
        // Write a line with an invalid expression – this will cause an inline error.
        builder.Writeln("Missing: <<[model.MissingProperty]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (simulating a real-world scenario).
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Build the report with InlineErrorMessages option enabled.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        Model model = new Model();

        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);

        // -----------------------------------------------------------------
        // 4. Validate the results.
        // -----------------------------------------------------------------
        // The success flag should be false because there is an invalid tag.
        Console.WriteLine($"Success flag returned by BuildReport: {success}");

        // The generated document should contain an inline error message.
        string docText = doc.GetText();
        bool containsError = docText.Contains("Error", StringComparison.OrdinalIgnoreCase);
        Console.WriteLine($"Document contains inline error message: {containsError}");

        // Simple assertion logic – in a real test framework you would use asserts.
        if (!success && containsError)
        {
            Console.WriteLine("InlineErrorMessages test passed.");
        }
        else
        {
            Console.WriteLine("InlineErrorMessages test failed.");
        }
    }
}
