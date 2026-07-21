using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Folder where the rendered PDF will be saved.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        // Simulated USB drive folder that contains TrueType fonts.
        string usbFontsDir = Path.Combine(Directory.GetCurrentDirectory(), "UsbFonts");

        // Ensure the directories exist.
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(usbFontsDir);

        // Create a sample document that contains special Unicode symbols.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial Unicode MS"; // Font with broad Unicode coverage (system font).
        builder.Writeln("Document with special symbols:");
        builder.Writeln("Greek: α β γ δ ε");
        builder.Writeln("Cyrillic: А Б В Г Д");
        builder.Writeln("Emoji: 😀 😃 😄");

        // Configure FontSettings to load fonts from the USB folder.
        FontSettings fontSettings = new FontSettings();
        // The second argument (true) enables recursive scanning of subfolders.
        fontSettings.SetFontsFolder(usbFontsDir, true);
        // Apply the FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        string pdfPath = Path.Combine(artifactsDir, "SpecialSymbols.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new Exception("PDF file was not created.");

        // Informative output (no user interaction required).
        Console.WriteLine($"PDF successfully saved to: {pdfPath}");
    }
}
