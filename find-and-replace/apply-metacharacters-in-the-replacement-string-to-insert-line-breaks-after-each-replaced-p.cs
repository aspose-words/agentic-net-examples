using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with placeholder text in several paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph contains [PLACE].");
        builder.Writeln("Second paragraph also has [PLACE] here.");
        builder.Writeln("Third paragraph does not have the token.");
        builder.Writeln("[PLACE] at the start of this paragraph.");
        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Replace the placeholder with a string that includes a paragraph break metacharacter (\r).
        // The \r inserts a new paragraph (line break) after the replacement text.
        const string findText = "[PLACE]";
        const string replaceText = "Inserted text\r";
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loaded.Range.Replace(findText, replaceText, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Optional: indicate success.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Modified document saved to '{outputPath}'.");
    }
}
