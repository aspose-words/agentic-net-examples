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
        // Define paths for the sample document and the simulated USB font folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        string usbFontDir = Path.Combine(Directory.GetCurrentDirectory(), "UsbFonts");
        string pdfPath = Path.Combine(outputDir, "SpecialSymbols.pdf");

        // Ensure output and USB font directories exist.
        Directory.CreateDirectory(outputDir);
        Directory.CreateDirectory(usbFontDir);

        // Attempt to copy a system TrueType font into the USB folder.
        // This makes the font available for Aspose.Words when it searches the USB directory.
        try
        {
            // Get the first system font folder reported by Aspose.Words.
            string systemFontFolder = SystemFontSource.GetSystemFontFolders().FirstOrDefault();
            if (!string.IsNullOrEmpty(systemFontFolder))
            {
                // Find any .ttf file in that folder.
                string fontFile = Directory.GetFiles(systemFontFolder, "*.ttf", SearchOption.TopDirectoryOnly).FirstOrDefault();
                if (!string.IsNullOrEmpty(fontFile))
                {
                    string destFile = Path.Combine(usbFontDir, Path.GetFileName(fontFile));
                    File.Copy(fontFile, destFile, overwrite: true);
                }
            }
        }
        catch
        {
            // If copying fails, continue – the example will still demonstrate font settings configuration.
        }

        // Create a new document and add text containing special Unicode symbols.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // Use a common font that we likely copied to the USB folder.
        builder.Writeln("Document with special symbols:");
        builder.Writeln("Greek: α β γ δ ε");
        builder.Writeln("Cyrillic: Д Ж З И Й");
        builder.Writeln("Chinese: 汉字测试");
        builder.Writeln("Emoji: 😀 😃 😄 😁");

        // Configure FontSettings to load fonts from the simulated USB drive.
        FontSettings fontSettings = new FontSettings();
        // The second parameter 'true' enables recursive search in subfolders.
        fontSettings.SetFontsFolder(usbFontDir, recursive: true);
        doc.FontSettings = fontSettings;

        // Save the document to PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use subsetting (default) to keep the PDF size reasonable.
            EmbedFullFonts = false
        };
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("The PDF file was not created.", pdfPath);

        // Optionally, output the location of the generated file.
        Console.WriteLine($"PDF saved to: {pdfPath}");
    }
}
