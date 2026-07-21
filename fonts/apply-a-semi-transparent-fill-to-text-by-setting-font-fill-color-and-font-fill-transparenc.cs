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

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This text has a semi‑transparent fill.");

        // Access the font of the recently added run.
        Aspose.Words.Font font = builder.Font;

        // Ensure the fill is solid.
        font.Fill.Solid();

        // Create an Aspose.Drawing.Color (blue with full opacity).
        Aspose.Drawing.Color aspColor = Aspose.Drawing.Color.FromArgb(255, 0, 0, 255);

        // Convert to System.Drawing.Color because Fill.Color expects that type.
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());

        // Set the fill color.
        font.Fill.Color = sysColor;

        // Set the transparency to 50% (0.5). 0.0 = opaque, 1.0 = clear.
        font.Fill.Transparency = 0.5;

        // Output the applied fill properties for verification.
        Console.WriteLine($"Fill type: {font.Fill.FillType}");
        Console.WriteLine($"Fill color: {font.Fill.Color}");
        Console.WriteLine($"Transparency: {font.Fill.Transparency}");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SemiTransparentFill.docx");
        doc.Save(outputPath);

        // Confirm that the file was created.
        Console.WriteLine(File.Exists(outputPath)
            ? $"Document saved successfully to: {outputPath}"
            : "Failed to save the document.");
    }
}
