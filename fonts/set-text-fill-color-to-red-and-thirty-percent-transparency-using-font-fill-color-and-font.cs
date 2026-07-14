using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // For Aspose.Drawing.Color

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with a single run of text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text with fill color and transparency.");

        // Retrieve the first run that was just added.
        Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

        // Access the Fill formatting of the run's font.
        Fill fill = run.Font.Fill;

        // Ensure the fill is a solid fill.
        fill.Solid();

        // Create an Aspose.Drawing.Color (red) and convert it to System.Drawing.Color.
        Aspose.Drawing.Color asposeRed = Aspose.Drawing.Color.Red;
        System.Drawing.Color sysRed = System.Drawing.Color.FromArgb(asposeRed.ToArgb());

        // Set the fill color to red (System.Drawing.Color required).
        fill.Color = sysRed;

        // Set the fill transparency to 30% (0.3).
        fill.Transparency = 0.3;

        // Optional validation (can be removed if not needed).
        Console.WriteLine($"Fill color ARGB: {fill.Color.ToArgb():X8}");
        Console.WriteLine($"Fill transparency: {fill.Transparency * 100}%");

        // Save the document to a file.
        string outputPath = "FontFillExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
