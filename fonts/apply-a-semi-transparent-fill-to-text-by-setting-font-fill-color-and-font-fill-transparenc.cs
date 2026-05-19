using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class SemiTransparentFillExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with a single run of text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This text has a semi‑transparent fill.");

        // Retrieve the run that was just added.
        Run run = doc.FirstSection.Body.FirstParagraph.Runs[0];

        // Ensure the fill type is solid before setting color and transparency.
        run.Font.Fill.Solid();

        // Create an Aspose.Drawing.Color (red) and convert it to System.Drawing.Color.
        Aspose.Drawing.Color asposeColor = Aspose.Drawing.Color.FromArgb(255, 255, 0, 0);
        run.Font.Fill.Color = System.Drawing.Color.FromArgb(asposeColor.ToArgb());

        // Set the fill transparency (0.0 = opaque, 1.0 = fully transparent).
        run.Font.Fill.Transparency = 0.5; // 50 % transparent

        // Optional: validate the fill properties.
        Console.WriteLine($"Fill Type: {run.Font.Fill.FillType}");
        Console.WriteLine($"Fill Color (ARGB): {run.Font.Fill.Color.ToArgb()}");
        Console.WriteLine($"Transparency: {run.Font.Fill.Transparency * 100}%");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SemiTransparentFill.docx");

        // Save the document to the specified path.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
