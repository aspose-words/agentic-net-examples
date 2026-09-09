using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for the example files.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeWordsFindReplaceExample");
        Directory.CreateDirectory(workFolder);

        // Define input and output file paths.
        string inputPath = Path.Combine(workFolder, "input.docx");
        string outputPath = Path.Combine(workFolder, "output.docx");

        // -------------------------------------------------
        // Create a sample document containing HTML tags.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(
            "This is a sample text with HTML tags: " +
            "<b>Bold</b>, <i>Italic</i>, and a link " +
            "<a href='https://example.com'>Example</a>.");
        doc.Save(inputPath);

        // -------------------------------------------------
        // Load the document we just created.
        // -------------------------------------------------
        Document loaded = new Document(inputPath);

        // -------------------------------------------------
        // Define a regular expression that matches any HTML tag.
        // -------------------------------------------------
        Regex htmlTagRegex = new Regex(@"<[^>]+>", RegexOptions.Compiled);

        // Perform the replacement: remove all HTML tags.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loaded.Range.Replace(htmlTagRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No HTML tags were found to replace.");

        // -------------------------------------------------
        // Save the modified document.
        // -------------------------------------------------
        loaded.Save(outputPath);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Replaced {replacedCount} HTML tag(s).");
        Console.WriteLine($"Modified document saved to: {outputPath}");
    }
}
