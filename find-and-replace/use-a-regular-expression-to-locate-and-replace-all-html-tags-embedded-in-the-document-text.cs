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
        builder.Writeln("This is a <b>bold</b> word.");
        builder.Writeln("Here is a <a href=\"https://example.com\">link</a>.");
        builder.Writeln("Plain text line.");

        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Regex to match any HTML tag.
        Regex htmlTagRegex = new Regex(@"<[^>]+>", RegexOptions.Compiled);

        // Replace all HTML tags with an empty string.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loaded.Range.Replace(htmlTagRegex, string.Empty, options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("No HTML tags were found to replace.");
        }

        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
