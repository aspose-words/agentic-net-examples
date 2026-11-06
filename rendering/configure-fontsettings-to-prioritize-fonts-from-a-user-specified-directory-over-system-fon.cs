using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define paths for the sample document and output PDF.
        string outputPdfPath = "RenderedWithCustomFonts.pdf";

        // Create a temporary folder that will hold user‑specified fonts.
        string customFontsFolder = Path.Combine(Directory.GetCurrentDirectory(), "CustomFonts");
        Directory.CreateDirectory(customFontsFolder);
        // NOTE: In a real scenario you would copy font files (e.g., .ttf, .otf) into this folder.
        // For this example we only demonstrate the configuration of FontSettings.

        // Create a simple document with a paragraph that specifies a font name.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "CustomFont"; // Font name that we expect to be resolved from the custom folder.
        builder.Writeln("This text should be rendered using a font from the custom fonts folder if available.");

        // Configure FontSettings to prioritize the custom fonts folder.
        FontSettings fontSettings = new FontSettings();
        // The second argument (true) enables recursive search in subfolders.
        fontSettings.SetFontsFolder(customFontsFolder, true);
        // Assign the FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF using the configured FontSettings.
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(outputPdfPath))
            throw new FileNotFoundException("The PDF file was not created.", outputPdfPath);

        // Optional: Output a simple confirmation (no user interaction required).
        Console.WriteLine($"PDF rendered successfully to '{outputPdfPath}'.");
    }
}
