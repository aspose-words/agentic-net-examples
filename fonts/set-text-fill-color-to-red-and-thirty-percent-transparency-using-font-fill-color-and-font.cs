using System;
using Aspose.Words;
using Aspose.Drawing;
using System.IO;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with some text.
        builder.Writeln("Hello, Aspose.Words!");

        // Retrieve the first run (the text we just added).
        Run run = doc.FirstSection.Body.FirstParagraph.Runs[0];
        Aspose.Words.Font font = run.Font; // Explicitly use Aspose.Words.Font

        // Create a red color using Aspose.Drawing.Color.
        Aspose.Drawing.Color asposeRed = Aspose.Drawing.Color.FromArgb(255, 255, 0, 0);
        // Convert to System.Drawing.Color because Font.Fill.Color expects it.
        System.Drawing.Color sysRed = System.Drawing.Color.FromArgb(asposeRed.ToArgb());

        // Set fill color and transparency.
        font.Fill.Color = sysRed;
        font.Fill.Transparency = 0.3f;

        // Validate that the properties were set correctly.
        bool colorMatches = font.Fill.Color.ToArgb() == sysRed.ToArgb();
        bool transparencyMatches = Math.Abs(font.Fill.Transparency - 0.3f) < 0.0001f;

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Ensure the output file exists.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");

        // Output validation results.
        Console.WriteLine($"Color set correctly: {colorMatches}");
        Console.WriteLine($"Transparency set correctly: {transparencyMatches}");
    }
}
