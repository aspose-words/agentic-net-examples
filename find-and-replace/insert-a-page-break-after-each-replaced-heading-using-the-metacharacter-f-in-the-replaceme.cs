using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample headings that we will replace.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.Writeln("Chapter 2");
        builder.Writeln("Chapter 3");

        // Perform find-and-replace: replace each heading text with the same text
        // followed by a page break using the form‑feed metacharacter (\f).
        // The replacement string contains the original text plus the page break.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = doc.Range.Replace("Chapter", "Chapter\f", options);

        // Validate that at least one replacement was made.
        if (replacedCount == 0)
            throw new InvalidOperationException("No headings were replaced.");

        // Save the modified document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ModifiedDocument.docx");
        doc.Save(outputPath);

        // Inform that the process completed successfully.
        Console.WriteLine($"Document saved to: {outputPath}");
        Console.WriteLine($"Number of replacements performed: {replacedCount}");
    }
}
