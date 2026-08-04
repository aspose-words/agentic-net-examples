using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for artifacts and the simulated USB drive.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string usbFontsDir = Path.Combine(artifactsDir, "UsbFonts");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(usbFontsDir);

        // Locate a system font file to copy to the USB folder.
        string systemFontFolder = SystemFontSource.GetSystemFontFolders().FirstOrDefault();
        if (string.IsNullOrEmpty(systemFontFolder))
            throw new InvalidOperationException("System font folder could not be located.");

        string sourceFontPath = Directory.GetFiles(systemFontFolder, "*.ttf").FirstOrDefault();
        if (string.IsNullOrEmpty(sourceFontPath))
            throw new InvalidOperationException("No TrueType font file found in the system font folder.");

        // Copy the font file to the USB folder.
        string fontFileName = Path.GetFileName(sourceFontPath);
        string usbFontPath = Path.Combine(usbFontsDir, fontFileName);
        File.Copy(sourceFontPath, usbFontPath, true);

        // Create a new document and write text containing special symbols.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = Path.GetFileNameWithoutExtension(fontFileName); // Use the copied font.
        builder.Writeln("Special symbols: ☺ ★ 漢字 🚀");

        // Configure FontSettings to load fonts from the USB folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(usbFontsDir, recursive: false);
        doc.FontSettings = fontSettings;

        // Save the document to PDF.
        string outputPath = Path.Combine(artifactsDir, "DocumentWithUsbFonts.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The PDF file was not created.", outputPath);
    }
}
