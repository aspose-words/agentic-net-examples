using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a few HTTP hyperlinks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample hyperlinks:");
        builder.InsertHyperlink("http://example.com", "http://example.com", false);
        builder.Writeln();
        builder.InsertHyperlink("http://test.com/page", "http://test.com/page", false);
        builder.Writeln();

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches the "http://" scheme.
        Regex httpPattern = new Regex(@"http://", RegexOptions.Compiled);

        // Configure find‑replace options to include fields (hyperlink field codes).
        FindReplaceOptions options = new FindReplaceOptions
        {
            // Ensure that the replace operation also processes field codes.
            // The default value is false, but we set it explicitly for clarity.
            IgnoreFields = false
        };

        // Perform the replacement, changing "http://" to "https://".
        int replacedCount = loaded.Range.Replace(httpPattern, "https://", options);

        // Validate that at least one hyperlink was updated.
        if (replacedCount == 0)
            throw new InvalidOperationException("No HTTP hyperlinks were found to replace.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
