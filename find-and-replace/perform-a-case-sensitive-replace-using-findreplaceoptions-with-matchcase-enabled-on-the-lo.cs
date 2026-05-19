using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with mixed‑case text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World. This is a Test. hello world.");

        // Save the document locally.
        const string inputPath = "sample.docx";
        doc.Save(inputPath);

        // Load the saved document.
        Document loaded = new Document(inputPath);

        // Configure case‑sensitive find‑replace options.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true
        };

        // Replace the exact case "Hello" with "Hi".
        int replacedCount = loaded.Range.Replace("Hello", "Hi", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "sample_output.docx";
        loaded.Save(outputPath);
    }
}
