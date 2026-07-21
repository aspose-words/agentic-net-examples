using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some paragraphs that contain the phrase we will replace.
        builder.Writeln("This is a sample document. The quick brown fox jumps over the lazy dog.");
        builder.Writeln("Another line with the phrase: TARGET_PHRASE appears here.");
        builder.Writeln("TARGET_PHRASE should be replaced everywhere.");

        // Phrase to search for and its replacement.
        string searchPhrase = "TARGET_PHRASE";
        string replacement = "REPLACED_TEXT";

        // Perform a simple case‑insensitive replace on the whole document range.
        int replacementsMade = doc.Range.Replace(searchPhrase, replacement);

        // Save the modified document.
        string outputFile = "ModifiedDocument.docx";
        doc.Save(outputFile);

        // Report the operation result.
        Console.WriteLine($"Replacements made: {replacementsMade}");
        Console.WriteLine($"Document saved to: {outputFile}");
    }
}
