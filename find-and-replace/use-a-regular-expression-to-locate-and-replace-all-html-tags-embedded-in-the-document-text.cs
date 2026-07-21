using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by Aspose.Words for drawing types
using Newtonsoft.Json; // Included as per package requirement

public class Program
{
    public static void Main()
    {
        // Create a sample document with HTML tags embedded in the text.
        const string inputPath = "sample.docx";
        const string outputPath = "cleaned.docx";

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a paragraph with <b>bold</b> HTML tag.");
        builder.Writeln("Another line with a <a href=\"https://example.com\">link</a>.");
        builder.Writeln("Plain text without tags.");
        doc.Save(inputPath);

        // Load the document we just created.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches any HTML tag.
        Regex htmlTagRegex = new Regex(@"<[^>]+>", RegexOptions.Compiled);

        // Perform the replacement: remove all HTML tags.
        int replacedCount = loaded.Range.Replace(htmlTagRegex, string.Empty, new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No HTML tags were found to replace.");

        // Save the cleaned document.
        loaded.Save(outputPath);

        // Optional: write a simple log to the console.
        Console.WriteLine($"Replaced {replacedCount} HTML tag(s). Output saved to '{outputPath}'.");
    }
}
