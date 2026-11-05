using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with HTTP hyperlinks.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("http://example.com/page1", "Example Page 1", false);
        builder.Writeln();
        builder.InsertHyperlink("http://test.org/resource", "Test Resource", false);
        builder.Writeln();
        builder.InsertHyperlink("http://sample.net", "Sample Site", false);
        // Save the input document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        var loadedDoc = new Document(inputPath);

        // Define a regex that captures the part of the URL after the scheme.
        var regex = new Regex(@"http://(?<rest>[^""\s]+)", RegexOptions.Compiled);

        // Replace the scheme with HTTPS while preserving the rest of the URL.
        var options = new FindReplaceOptions();
        int replacedCount = loadedDoc.Range.Replace(regex, "https://${rest}", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No HTTP URLs were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
