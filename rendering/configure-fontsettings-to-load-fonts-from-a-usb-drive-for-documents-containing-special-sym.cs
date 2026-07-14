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
        // Prepare output directories.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Simulate a USB drive by creating a folder that will contain font files.
        string usbFontsDir = Path.Combine(artifactsDir, "USBFonts");
        Directory.CreateDirectory(usbFontsDir);

        // Try to copy a TrueType font from the system font folders into the USB folder.
        // This ensures that the folder actually contains a font that Aspose.Words can use.
        bool fontCopied = false;
        foreach (string systemFolder in SystemFontSource.GetSystemFontFolders())
        {
            if (!Directory.Exists(systemFolder))
                continue;

            string ttfFile = Directory.EnumerateFiles(systemFolder, "*.ttf", SearchOption.TopDirectoryOnly).FirstOrDefault();
            if (ttfFile != null)
            {
                string destFile = Path.Combine(usbFontsDir, Path.GetFileName(ttfFile));
                File.Copy(ttfFile, destFile, overwrite: true);
                fontCopied = true;
                break;
            }
        }

        // Create a simple document that contains special Unicode symbols.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // Use a common font name; actual glyphs will be resolved via the USB font folder.
        builder.Writeln("Document with special symbols:");
        builder.Writeln("Greek Omega: Ω");
        builder.Writeln("Chinese characters: 漢字");
        builder.Writeln("Emoji: 😊");

        // Configure FontSettings to load fonts from the simulated USB drive.
        FontSettings fontSettings = new FontSettings();
        // The second parameter 'recursive' is set to true to include subfolders if any.
        fontSettings.SetFontsFolder(usbFontsDir, recursive: true);
        doc.FontSettings = fontSettings;

        // Save the document to PDF.
        string pdfPath = Path.Combine(artifactsDir, "DocumentWithSpecialSymbols.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF output.");

        // Optional: simple check that the PDF contains a font embedding marker.
        // This is a compile‑safe way to verify that fonts were considered during rendering.
        string pdfContent = File.ReadAllText(pdfPath);
        bool containsFontMarker = pdfContent.Contains("/FontFile") || pdfContent.Contains("/FontFile2") || pdfContent.Contains("/FontFile3");
        if (!containsFontMarker)
            Console.WriteLine("Warning: No embedded font markers were found in the PDF.");

        Console.WriteLine($"PDF successfully saved to: {pdfPath}");
    }
}
