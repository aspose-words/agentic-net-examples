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

        // Add left‑to‑right text.
        builder.Writeln("Hello world!");

        // Add right‑to‑left paragraphs.
        builder.ParagraphFormat.Bidi = true;
        builder.Writeln("שלום עולם!");      // Hebrew
        builder.Writeln("مرحبا بالعالم!");   // Arabic

        // Configure save options to add BiDi marks.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            Encoding = Encoding.Unicode,
            AddBidiMarks = true
        };

        // Save the document as plain text.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BidiMarks.txt");
        doc.Save(outputPath, saveOptions);

        // Read and display the saved text.
        string savedText = File.ReadAllText(outputPath, Encoding.Unicode);
        Console.WriteLine("Saved text with BiDi marks:");
        Console.WriteLine(savedText);
    }
}
