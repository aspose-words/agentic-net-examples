using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json; // Required package reference

public class HtmlTagRemover
{
    public static void Main()
    {
        // Step 1: Create a sample document with embedded HTML tags.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a <b>bold</b> word, an <i>italic</i> word, and a <a href='https://example.com'>link</a>.");
        builder.Writeln("Another line with <span style='color:red'>colored text</span>.");

        // Save the original document.
        const string inputPath = "input.docx";
        sampleDoc.Save(inputPath);

        // Step 2: Load the document from the file system.
        Document loadedDoc = new Document(inputPath);

        // Step 3: Define a regular expression that matches any HTML tag.
        Regex htmlTagRegex = new Regex(@"<[^>]+>", RegexOptions.Compiled);

        // Step 4: Perform the replacement using Aspose.Words Range.Replace.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loadedDoc.Range.Replace(htmlTagRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No HTML tags were found to replace.");

        // Step 5: Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Inform the user of the result.
        Console.WriteLine($"Replaced {replacedCount} HTML tag(s). Output saved to '{outputPath}'.");
    }
}
