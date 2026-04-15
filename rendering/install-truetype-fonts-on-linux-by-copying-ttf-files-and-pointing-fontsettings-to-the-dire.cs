using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define directories for artifacts and fonts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string fontsDir = Path.Combine(artifactsDir, "Fonts");

        // Ensure the directories exist.
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(fontsDir);

        // Create a dummy TrueType font file in the fonts directory.
        // In a real scenario, copy actual .ttf files here.
        string dummyFontPath = Path.Combine(fontsDir, "DummyFont.ttf");
        if (!File.Exists(dummyFontPath))
        {
            // Write a minimal TTF header (not a valid font, but serves as a placeholder).
            byte[] minimalTtfHeader = new byte[]
            {
                0x00,0x01,0x00,0x00, // sfnt version
                0x00,0x00,0x00,0x00, // numTables
                0x00,0x00,0x00,0x00, // searchRange
                0x00,0x00,0x00,0x00, // entrySelector
                0x00,0x00,0x00,0x00  // rangeShift
            };
            File.WriteAllBytes(dummyFontPath, minimalTtfHeader);
        }

        // Create a new document and add some text using the custom font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "DummyFont";
        builder.Writeln("This text should be rendered with the custom DummyFont.");

        // Configure FontSettings to use the fonts folder we created.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(fontsDir, recursive: false);
        doc.FontSettings = fontSettings;

        // Save the document to PDF.
        string outputPdfPath = Path.Combine(artifactsDir, "Output.pdf");
        doc.Save(outputPdfPath);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdfPath))
            throw new Exception("Failed to create the output PDF file.");

        // Optionally, indicate success (no interactive prompts required).
        Console.WriteLine("Document rendered and saved to: " + outputPdfPath);
    }
}
