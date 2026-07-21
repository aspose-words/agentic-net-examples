using System;
using System.IO;
using System.Text;
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
        builder.Writeln("שלום עולם!"); // Hebrew
        builder.Writeln("مرحبا بالعالم!"); // Arabic

        // Configure TxtSaveOptions to add BiDi marks.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            Encoding = Encoding.Unicode,
            AddBidiMarks = true
        };

        // Save the document as plain text with BiDi marks.
        string outputFile = Path.Combine(Environment.CurrentDirectory, "BidiMarks.txt");
        doc.Save(outputFile, saveOptions);

        // Demonstrate that the file was saved (no user interaction required).
        string content = File.ReadAllText(outputFile, Encoding.Unicode);
        Console.WriteLine($"Saved file length: {content.Length} characters");
    }
}
