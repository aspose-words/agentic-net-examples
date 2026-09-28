using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with HTTP URLs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Visit http://example.com for examples.");
        builder.Writeln("Another link: http://test.org/page?query=1.");
        builder.Writeln("Secure site already: https://secure.com should stay unchanged.");
        // Save the input document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regex that matches the "http://" scheme.
        Regex httpSchemeRegex = new Regex(@"http://(?=\S)", RegexOptions.Compiled);

        // Perform the replacement, changing "http://" to "https://".
        int replacedCount = loaded.Range.Replace(httpSchemeRegex, "https://", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one URL scheme replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Optional: write a simple confirmation to the console.
        Console.WriteLine($"Replaced {replacedCount} URL scheme(s). Output saved to '{outputPath}'.");
    }
}
