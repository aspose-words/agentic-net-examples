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

        // Mark the next paragraph as right‑to‑left.
        builder.ParagraphFormat.Bidi = true;
        builder.Writeln("שלום עולם!");   // Hebrew
        builder.Writeln("مرحبا بالعالم!"); // Arabic

        // Configure save options to add BiDi marks when exporting to plain text.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            Encoding = Encoding.Unicode,
            AddBidiMarks = true
        };

        // Determine an output path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.txt");

        // Save the document as plain text using the configured options.
        doc.Save(outputPath, saveOptions);

        // Optionally display a confirmation.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
