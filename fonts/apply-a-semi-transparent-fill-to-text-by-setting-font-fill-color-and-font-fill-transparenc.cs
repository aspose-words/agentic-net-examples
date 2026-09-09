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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line of text.
        builder.Writeln("Hello, semi‑transparent fill!");

        // Access the font of the last inserted run.
        Aspose.Words.Font font = builder.Font;

        // Define a solid fill color (red) using Aspose.Drawing.Color.
        Aspose.Drawing.Color fillColor = Aspose.Drawing.Color.Red;

        // Apply the fill color to the font. The Solid method expects System.Drawing.Color,
        // so convert the Aspose.Drawing.Color to System.Drawing.Color.
        font.Fill.Solid(System.Drawing.Color.FromArgb(fillColor.ToArgb()));

        // Set the fill transparency to 50% (0.5).
        font.Fill.Transparency = 0.5;

        // Validate that the properties were set correctly.
        if (font.Fill.Color.ToArgb() != fillColor.ToArgb() ||
            Math.Abs(font.Fill.Transparency - 0.5) > 0.0001)
        {
            throw new InvalidOperationException("Fill properties were not applied as expected.");
        }

        // Ensure the output directory exists.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "SemiTransparentFill.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved.", outputPath);
    }
}
