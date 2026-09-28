using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing HTML tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a <b>bold</b> word and a <a href=\"https://example.com\">link</a>.");
        builder.Writeln("Another line with <i>italic</i> text.");
        // Save the initial document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches HTML tags.
        Regex htmlTagRegex = new Regex(@"<[^>]+>", RegexOptions.Compiled);

        // Replace all HTML tags with an empty string.
        int replacedCount = loaded.Range.Replace(htmlTagRegex, string.Empty, new FindReplaceOptions());

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No HTML tags were found to replace.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
