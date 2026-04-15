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

        // Add a right‑to‑left Hebrew paragraph.
        builder.ParagraphFormat.Bidi = true;
        builder.Writeln("שלום עולם!");

        // Add a right‑to‑left Arabic paragraph.
        builder.Writeln("مرحبا بالعالم!");

        // Configure save options to add bi‑directional marks.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            Encoding = Encoding.Unicode,
            AddBidiMarks = true
        };

        // Save the document as plain text.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BiDiOutput.txt");
        doc.Save(outputPath, saveOptions);

        // Optional: display the saved text length to confirm execution.
        string savedText = File.ReadAllText(outputPath, Encoding.Unicode);
        Console.WriteLine($"Saved text length: {savedText.Length}");
    }
}
