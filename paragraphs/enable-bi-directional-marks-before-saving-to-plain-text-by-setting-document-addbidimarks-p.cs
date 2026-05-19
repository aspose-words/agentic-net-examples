using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a left‑to‑right paragraph.
        builder.Writeln("Hello world!");

        // Add right‑to‑left paragraphs.
        builder.ParagraphFormat.Bidi = true;
        builder.Writeln("שלום עולם!");          // Hebrew
        builder.Writeln("مرحبا بالعالم!");      // Arabic

        // Configure save options to insert BiDi marks before each RTL run.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            Encoding = System.Text.Encoding.Unicode,
            AddBidiMarks = true
        };

        // Save the document as plain text.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BidiText.txt");
        doc.Save(outputPath, saveOptions);
    }
}
