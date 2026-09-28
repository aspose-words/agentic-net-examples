using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define paths for the sample document and output PDF.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, "RenderedDocument.pdf");

        // Create a folder that will hold user‑specified fonts.
        string customFontsFolder = Path.Combine(Directory.GetCurrentDirectory(), "CustomFonts");
        Directory.CreateDirectory(customFontsFolder);
        // (In a real scenario you would copy font files into this folder.)

        // Create a simple document that uses a specific font name.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "CustomFont"; // Font name that we expect to resolve from the custom folder.
        builder.Writeln("This text should be rendered using a font from the custom fonts folder if available.");

        // Configure FontSettings to prioritize the custom fonts folder.
        FontSettings fontSettings = new FontSettings();
        // The second argument (false) indicates that subfolders are not searched.
        // Adding the custom folder first gives it higher priority over system fonts.
        fontSettings.SetFontsFolder(customFontsFolder, false);
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF rendering failed; output file not found.");

        // Optionally, output the path of the generated file.
        Console.WriteLine($"PDF rendered successfully to: {pdfPath}");
    }
}
