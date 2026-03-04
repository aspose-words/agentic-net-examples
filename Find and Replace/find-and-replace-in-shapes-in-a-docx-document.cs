using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceInShapes
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Create FindReplaceOptions. By default IgnoreShapes = false,
        // which means the replace operation will also process text inside shapes.
        FindReplaceOptions options = new FindReplaceOptions();

        // Replace the target text with the new text throughout the document,
        // including any text that resides inside shapes (e.g., text boxes, WordArt).
        int replacements = doc.Range.Replace("OldText", "NewText", options);

        Console.WriteLine($"Number of replacements made: {replacements}");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
