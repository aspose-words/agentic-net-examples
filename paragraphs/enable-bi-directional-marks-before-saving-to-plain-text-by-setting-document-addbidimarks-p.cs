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

        // Mark the next paragraph as right‑to‑left.
        builder.ParagraphFormat.Bidi = true;
        builder.Writeln("שלום עולם!"); // Hebrew
        builder.Writeln("مرحبا بالعالم!"); // Arabic

        // Prepare save options for plain‑text output.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            Encoding = Encoding.Unicode,
            AddBidiMarks = true // Enable bi‑directional marks.
        };

        // Ensure the output folder exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document as a .txt file with the specified options.
        string txtPath = Path.Combine(outputDir, "BidiMarks.txt");
        doc.Save(txtPath, saveOptions);

        // Read the saved file and display its contents (for demonstration purposes).
        string savedText = File.ReadAllText(txtPath, Encoding.Unicode);
        Console.WriteLine("Saved text content:");
        Console.WriteLine(savedText);
    }
}
