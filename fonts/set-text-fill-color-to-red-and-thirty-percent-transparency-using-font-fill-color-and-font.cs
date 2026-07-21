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

        // Use DocumentBuilder to add a paragraph with a run of text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text with custom fill color and transparency.");

        // Get the Font of the last run (the one we just added).
        // The Run is the first (and only) run of the first paragraph.
        Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];
        Aspose.Words.Font font = run.Font; // Fully qualified to avoid ambiguity

        // Ensure the fill is set to a solid color.
        font.Fill.Solid();

        // Create a red color using Aspose.Drawing.Color.
        Aspose.Drawing.Color asposeRed = Aspose.Drawing.Color.Red;

        // Convert Aspose.Drawing.Color to System.Drawing.Color because Fill.Color expects System.Drawing.Color.
        System.Drawing.Color sysRed = System.Drawing.Color.FromArgb(asposeRed.ToArgb());

        // Apply the color and set transparency to 30% (0.3).
        font.Fill.Color = sysRed;
        font.Fill.Transparency = 0.3; // 30% transparent

        // Validate that the properties were set correctly.
        if (font.Fill.Transparency != 0.3 ||
            font.Fill.Color.ToArgb() != sysRed.ToArgb())
        {
            throw new InvalidOperationException("Font fill properties were not set as expected.");
        }

        // Define output path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "FontFillExample.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
        }
    }
}
