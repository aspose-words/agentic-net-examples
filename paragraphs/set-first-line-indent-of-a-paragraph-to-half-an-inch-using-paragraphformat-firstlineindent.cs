using System;
using System.IO;
using Aspose.Words;

public class SetFirstLineIndentExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and set formatting.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set first line indent to half an inch (36 points).
        builder.ParagraphFormat.FirstLineIndent = 36.0;

        // Write a sample paragraph.
        builder.Writeln("This paragraph has a first line indent of half an inch.");

        // Determine an output path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FirstLineIndent.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
