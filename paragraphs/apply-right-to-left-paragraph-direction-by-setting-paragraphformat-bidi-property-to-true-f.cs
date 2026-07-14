using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Arabic sample text.
        string arabicText = "مرحبا بالعالم! هذا نص عربي.";

        // Apply right‑to‑left direction to the current paragraph.
        builder.ParagraphFormat.Bidi = true;

        // Write the Arabic text into the paragraph.
        builder.Writeln(arabicText);

        // Save the document to a file.
        string outputPath = "BidiParagraph.docx";
        doc.Save(outputPath);
    }
}
