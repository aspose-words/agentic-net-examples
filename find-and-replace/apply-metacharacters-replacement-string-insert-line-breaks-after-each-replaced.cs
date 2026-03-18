using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class InsertParagraphBreakAfterReplace
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs that contain the word we will replace.
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");

        // Use the Range.Replace method with the meta‑character &p.
        // The pattern "Paragraph" will be replaced by "Paragraph&p",
        // which inserts a paragraph break after each occurrence.
        FindReplaceOptions options = new FindReplaceOptions(); // default options
        int replacementsMade = doc.Range.Replace("Paragraph", "Paragraph&p", options);

        // Output the number of replacements for verification.
        Console.WriteLine($"Replacements performed: {replacementsMade}");

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
