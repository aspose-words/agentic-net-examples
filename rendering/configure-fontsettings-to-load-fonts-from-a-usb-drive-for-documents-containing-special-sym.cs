using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Simulate a USB drive folder that contains fonts.
        string usbFontsFolder = Path.Combine(Path.GetTempPath(), "UsbFonts");
        Directory.CreateDirectory(usbFontsFolder);

        // Copy a common system font to the simulated USB folder (if it exists).
        // This ensures the document can find a font that supports special symbols.
        string systemFontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf");
        if (File.Exists(systemFontPath))
        {
            string destFontPath = Path.Combine(usbFontsFolder, "arial.ttf");
            File.Copy(systemFontPath, destFontPath, true);
        }

        // Configure FontSettings to load fonts from the USB folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(usbFontsFolder, false);

        // Create a sample document that contains special Unicode symbols.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // Use the font we placed on the USB drive.
        builder.Writeln("Special symbols: Ω ∑ ∞ 𝄞");

        // Apply the FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        string outputPath = Path.Combine(Path.GetTempPath(), "RenderedDocument.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created successfully.
        if (!File.Exists(outputPath))
            throw new Exception("PDF rendering failed; output file not found.");

        // (Optional) Clean up the simulated USB folder.
        // Directory.Delete(usbFontsFolder, true);
    }
}
