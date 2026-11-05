using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing HTML tags within the text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph with <b>bold</b> HTML tag.");
        builder.Writeln("Another line with <a href=\"https://example.com\">link</a> and <img src=\"image.png\"/>.");
        // Save the initial document to a local file.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document that will be processed.
        Document loadedDoc = new Document(inputPath);

        // Regular expression that matches any HTML tag, e.g., <tag ...>.
        Regex htmlTagRegex = new Regex(@"<[^>]+>", RegexOptions.Compiled);

        // Replace all HTML tags with an empty string (remove them).
        int replacedCount = loadedDoc.Range.Replace(htmlTagRegex, string.Empty, new FindReplaceOptions());

        // Ensure that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("No HTML tags were found to replace.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
