using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define paths for the sample document and the output PDF.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPdfPath = Path.Combine(artifactsDir, "DocumentWithSpecialSymbols.pdf");

        // Simulate a USB drive by creating a folder that will contain custom fonts.
        // In a real scenario this folder would be the mount point of the USB drive.
        string usbFontsPath = Path.Combine(Directory.GetCurrentDirectory(), "UsbFonts");
        Directory.CreateDirectory(usbFontsPath);
        // (Optional) Copy a font file into the folder if you have one available.
        // For this example we leave the folder empty to demonstrate fallback behavior.

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text that contains special Unicode symbols.
        builder.Font.Name = "NonExistentFont"; // Intentionally use a font that is not installed.
        builder.Writeln("Hello, world! こんにちは世界 🌍 🚀");

        // Configure FontSettings to load fonts from the simulated USB drive.
        FontSettings fontSettings = new FontSettings();
        // The second argument 'true' enables recursive search in subfolders.
        fontSettings.SetFontsFolder(usbFontsPath, recursive: true);
        // Assign the FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Save the document to PDF. During rendering Aspose.Words will use the fonts
        // from the specified folder, falling back to system fonts where necessary.
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Clean up (optional): delete the temporary folders and files.
        // Directory.Delete(usbFontsPath, recursive: true);
        // Directory.Delete(artifactsDir, recursive: true);
    }
}
