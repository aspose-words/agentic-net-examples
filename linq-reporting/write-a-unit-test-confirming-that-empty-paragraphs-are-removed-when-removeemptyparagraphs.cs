using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Property that will be empty in the report.
    public string Empty { get; set; } = string.Empty;

    // Sample non‑empty property.
    public string Name { get; set; } = "John Doe";
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph with a normal value.
        builder.Writeln("Name: <<[model.Name]>>");

        // Paragraph that contains only a tag which resolves to an empty string.
        builder.Writeln("<<[model.Empty]>>");

        // -----------------------------------------------------------------
        // 2. Build the report with the RemoveEmptyParagraphs option enabled.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // The root object name in the template is "model".
        engine.BuildReport(doc, new Model(), "model");

        // -----------------------------------------------------------------
        // 3. Verify that the empty paragraph has been removed.
        // -----------------------------------------------------------------
        // Get the document text; paragraphs are separated by '\r'.
        string fullText = doc.GetText();
        string[] paragraphs = fullText.Split('\r');

        // Count paragraphs that contain visible text.
        int nonEmptyParagraphCount = 0;
        foreach (string p in paragraphs)
        {
            if (!string.IsNullOrWhiteSpace(p))
                nonEmptyParagraphCount++;
        }

        // The template had two paragraphs, but the second one should be gone.
        if (nonEmptyParagraphCount == 1)
        {
            Console.WriteLine("Test passed: empty paragraphs were removed.");
        }
        else
        {
            Console.WriteLine("Test failed: empty paragraph was not removed.");
            Environment.Exit(1);
        }
    }
}
