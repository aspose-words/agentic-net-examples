using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Arabic sample text.
        string arabicText = "مرحبا بالعالم!";

        // Set the paragraph direction to right‑to‑left.
        builder.ParagraphFormat.Bidi = true;

        // Write the Arabic text; it will be laid out RTL.
        builder.Writeln(arabicText);

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "BidiParagraph.docx");
        doc.Save(outputPath);
    }
}
