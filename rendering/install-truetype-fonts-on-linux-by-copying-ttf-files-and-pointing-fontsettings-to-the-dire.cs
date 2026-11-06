using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define a folder to hold TrueType fonts.
        string fontsFolder = Path.Combine(Directory.GetCurrentDirectory(), "fonts");
        Directory.CreateDirectory(fontsFolder);

        // Path for a sample TrueType font file.
        string fontFilePath = Path.Combine(fontsFolder, "SampleFont.ttf");

        // Write a minimal TrueType font file (dummy content) to the folder.
        // In a real scenario, copy actual .ttf files here.
        byte[] dummyFontData = new byte[] { 0x00, 0x01, 0x00, 0x00, 0x00, 0x0C, 0x00, 0x80 };
        File.WriteAllBytes(fontFilePath, dummyFontData);

        // Configure FontSettings to use the custom fonts folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(fontsFolder, false);

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "SampleFont"; // Attempt to use the custom font.
        builder.Writeln("This text is rendered using the custom TrueType font if available.");

        // Assign the FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Save the document as PDF.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The PDF output file was not created.", outputPath);

        // Simple confirmation (no interactive prompts).
        Console.WriteLine("Document rendered and saved to: " + outputPath);
    }
}
