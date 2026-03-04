using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceWithHighlight
{
    static void Main()
    {
        // Create a new document and add some sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("The quick brown fox is swift.");

        // Configure find/replace options.
        FindReplaceOptions options = new FindReplaceOptions();

        // Highlight the replaced text with a light yellow background.
        options.ApplyFont.HighlightColor = Color.LightYellow;

        // Perform the replace operation.
        // All occurrences of the word "quick" will be replaced with "fast" and highlighted.
        int replacements = doc.Range.Replace("quick", "fast", options);

        Console.WriteLine($"Number of replacements made: {replacements}");

        // Save the modified document.
        doc.Save("FindReplaceHighlighted.docx");
    }
}
