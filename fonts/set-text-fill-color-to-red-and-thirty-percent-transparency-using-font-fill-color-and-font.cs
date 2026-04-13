using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with some text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Text with red fill color and 30% transparency.");

        // Retrieve the first run of the first paragraph.
        Run run = (Run)doc.FirstSection.Body.FirstParagraph.Runs[0];

        // Access the font of the run.
        Aspose.Words.Font font = run.Font;

        // Access the fill formatting of the font.
        Aspose.Words.Drawing.Fill fill = font.Fill;

        // Ensure the fill is a solid fill.
        fill.Solid();

        // Create an Aspose.Drawing.Color (red) and convert it to System.Drawing.Color.
        Aspose.Drawing.Color aspRed = Aspose.Drawing.Color.Red;
        System.Drawing.Color sysRed = System.Drawing.Color.FromArgb(aspRed.ToArgb());

        // Set the fill color to red (System.Drawing.Color).
        fill.Color = sysRed;

        // Set the fill transparency to 30% (0.3).
        fill.Transparency = 0.3;

        // Validate that the properties were set correctly.
        if (fill.Color.ToArgb() != sysRed.ToArgb())
            throw new InvalidOperationException("Fill color was not set to red.");

        if (Math.Abs(fill.Transparency - 0.3) > 0.0001)
            throw new InvalidOperationException("Fill transparency was not set to 30%.");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FillColorTransparency.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
