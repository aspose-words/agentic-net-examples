using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // for Aspose.Drawing.Color

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure the fill type is solid.
        builder.Font.Fill.Solid();

        // Create an Aspose.Drawing.Color (red) and convert it to System.Drawing.Color.
        Aspose.Drawing.Color asposeRed = Aspose.Drawing.Color.Red;
        System.Drawing.Color sysRed = System.Drawing.Color.FromArgb(asposeRed.ToArgb());

        // Set the fill color to red.
        builder.Font.Fill.Color = sysRed;

        // Set the fill transparency to 30% (0.3 = 30% transparent, 0 = opaque).
        builder.Font.Fill.Transparency = 0.3;

        // Write a sample line that will use the configured fill.
        builder.Writeln("Sample text with red fill and 30% transparency.");

        // Validate that the properties were applied correctly.
        Aspose.Words.Font font = builder.Font;
        if (font.Fill.Color.ToArgb() != sysRed.ToArgb() ||
            Math.Abs(font.Fill.Transparency - 0.3) > 0.0001)
        {
            throw new Exception("Font fill properties were not set as expected.");
        }

        // Save the document to a file.
        const string outputPath = "FontFillExample.docx";
        doc.Save(outputPath);
    }
}
