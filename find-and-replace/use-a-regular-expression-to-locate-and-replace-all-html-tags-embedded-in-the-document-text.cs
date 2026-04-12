using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;

public class HtmlTagRemover
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert sample text that contains HTML tags.
        builder.Writeln("This is a <b>bold</b> word.");
        builder.Writeln("Here is a <a href=\"https://example.com\">link</a>.");
        builder.Writeln("An image tag: <img src=\"image.png\" alt=\"Sample\"/>.");
        builder.Writeln("Plain text without tags.");

        // Define a regular expression that matches any HTML tag.
        Regex htmlTagPattern = new Regex("<[^>]+>", RegexOptions.Compiled);

        // Replace all HTML tags with an empty string.
        int replacementCount = doc.Range.Replace(htmlTagPattern, string.Empty);

        // Validate that at least one replacement was made.
        if (replacementCount == 0)
            throw new InvalidOperationException("No HTML tags were found to replace.");

        // Save the modified document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Optional: write a simple confirmation to the console.
        Console.WriteLine($"Replaced {replacementCount} HTML tag(s). Output saved to: {outputPath}");
    }
}
