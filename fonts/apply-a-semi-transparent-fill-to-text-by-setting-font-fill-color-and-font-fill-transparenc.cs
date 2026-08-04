using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with a run of text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This text has a semi‑transparent fill.");

        // Access the Font of the last run (the one just added).
        Aspose.Words.Font font = builder.Font;

        // Ensure the fill is a solid fill.
        font.Fill.Solid();

        // Create an Aspose.Drawing.Color and convert it to System.Drawing.Color for the Fill.
        Aspose.Drawing.Color aspColor = Aspose.Drawing.Color.Blue;
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());

        // Set the fill color and transparency.
        font.Fill.Color = sysColor;
        font.Fill.Transparency = 0.5;

        // Validate the applied properties.
        if (Math.Abs(font.Fill.Transparency - 0.5) > 0.0001 ||
            font.Fill.Color.ToArgb() != sysColor.ToArgb())
        {
            throw new InvalidOperationException("Font fill properties were not set correctly.");
        }

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "SemiTransparentFill.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
        }
    }
}
