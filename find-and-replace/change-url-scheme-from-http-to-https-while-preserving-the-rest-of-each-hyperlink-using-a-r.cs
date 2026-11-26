using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a few hyperlinks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Sample document with hyperlinks:");
        builder.InsertHyperlink("http://example.com", "Example Site", false);
        builder.Writeln();
        builder.InsertHyperlink("http://contoso.com/page", "Contoso Page", false);
        builder.Writeln();

        // Save the input document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches the HTTP scheme.
        Regex httpRegex = new Regex(@"http://", RegexOptions.Compiled);

        // Replace all occurrences of "http://" with "https://".
        int replacedCount = loaded.Range.Replace(httpRegex, "https://", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No hyperlink URLs were updated.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
