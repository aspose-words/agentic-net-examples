using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder that will act as the Linux font directory.
        string fontsDir = Path.Combine(Directory.GetCurrentDirectory(), "fonts");
        Directory.CreateDirectory(fontsDir);

        // Create a dummy TrueType font file in the folder.
        // In a real scenario this would be a valid .ttf file copied from somewhere.
        string dummyFontPath = Path.Combine(fontsDir, "CustomFont.ttf");
        byte[] dummyFontData = new byte[] { 0x00, 0x01, 0x00, 0x00, 0x00, 0x0C, 0x00, 0x80, 0x00, 0x03, 0x00, 0x50 };
        File.WriteAllBytes(dummyFontPath, dummyFontData);

        // Point Aspose.Words to the folder that contains the fonts.
        FontSettings.DefaultInstance.SetFontsFolder(fontsDir, recursive: true);

        // Build a simple document that uses the custom font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "CustomFont"; // Name matches the dummy font file.
        builder.Writeln("This text is intended to use the custom TrueType font.");

        // Render the document to PDF.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the PDF output file.");

        Console.WriteLine($"PDF successfully saved to: {outputPath}");
    }
}
