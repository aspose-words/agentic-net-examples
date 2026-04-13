using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define output folder and ensure it exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "SemiTransparentFill.docx");

        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with a run of text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello with semi‑transparent fill!");

        // Retrieve the first run that was just added.
        Run run = doc.FirstSection.Body.FirstParagraph.Runs[0];
        Aspose.Words.Font font = run.Font;

        // Access the fill formatting of the font.
        Fill fill = font.Fill;

        // Set the fill color to blue using Aspose.Drawing.Color,
        // then convert it to System.Drawing.Color as required by the Fill API.
        fill.Color = System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Blue.ToArgb());

        // Set the fill transparency to 50% (0.5).
        fill.Transparency = 0.5;

        // Ensure the fill type is solid.
        fill.Solid();

        // Validate that the properties were applied correctly.
        if (fill.Color.ToArgb() != System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Blue.ToArgb()).ToArgb())
            throw new InvalidOperationException("Fill color was not set correctly.");

        if (Math.Abs(fill.Transparency - 0.5) > 0.0001)
            throw new InvalidOperationException("Fill transparency was not set correctly.");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
