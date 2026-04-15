using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder for easy content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph direction to right‑to‑left.
        builder.ParagraphFormat.Bidi = true;

        // Write an Arabic sentence. The paragraph will inherit the Bidi setting.
        builder.Writeln("مرحبا بالعالم!");

        // Define an output path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ArabicBidiParagraph.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
