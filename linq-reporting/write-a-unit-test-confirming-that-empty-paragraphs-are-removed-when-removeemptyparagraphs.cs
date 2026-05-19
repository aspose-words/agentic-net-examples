using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model for the report.
    public class ReportModel
    {
        // This property will be empty, causing an empty paragraph after the tag is processed.
        public string EmptyValue { get; set; } = string.Empty;

        // Regular property with a value.
        public string Name { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // Prepare file paths.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string templatePath = Path.Combine(artifactsDir, "Template.docx");
        string outputPath = Path.Combine(artifactsDir, "Result.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Paragraph with a normal value.
        builder.Writeln("<<[model.Name]>>");
        // Paragraph that will become empty after the engine processes it.
        builder.Writeln("<<[model.EmptyValue]>>");
        // Another paragraph to ensure the document still has content after removal.
        builder.Writeln("End of document.");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (required by the workflow rules).
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data source.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Name = "John Doe",
            EmptyValue = string.Empty // Explicitly empty.
        };

        // -----------------------------------------------------------------
        // 4. Configure the ReportingEngine with RemoveEmptyParagraphs option.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report. The root object name must match the tag prefix ("model").
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated document.
        loadedTemplate.Save(outputPath);

        // -----------------------------------------------------------------
        // 5. Verify that the empty paragraph was removed.
        // -----------------------------------------------------------------
        string resultText = loadedTemplate.GetText();

        // In Aspose.Words, a paragraph break is represented by '\r'.
        // Two consecutive '\r' indicate an empty paragraph between content.
        bool containsEmptyParagraph = resultText.Contains("\r\r");

        if (!containsEmptyParagraph)
        {
            Console.WriteLine("Test passed: empty paragraphs were removed.");
        }
        else
        {
            Console.WriteLine("Test failed: empty paragraph still present.");
        }

        // Optional: display the resulting text for manual inspection.
        Console.WriteLine("Resulting document text:");
        Console.WriteLine(resultText);
    }
}
