using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define output paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, "CustomFontLookup.pdf");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text using a font that is unlikely to be installed.
        // This will force Aspose.Words to use the custom FontSettings for lookup.
        builder.Font.Name = "NonExistentFont";
        builder.Writeln("This text should be rendered with a fallback font using custom FontSettings.");

        // Create a FontSettings instance.
        FontSettings fontSettings = new FontSettings();

        // (Optional) Specify a folder that contains custom fonts.
        // Here we create a subfolder "CustomFonts". It may be empty; Aspose.Words will still use it as a source.
        string customFontsFolder = Path.Combine(Directory.GetCurrentDirectory(), "CustomFonts");
        Directory.CreateDirectory(customFontsFolder);
        // Example of adding a folder as a font source (recursive = true).
        fontSettings.SetFontsFolder(customFontsFolder, recursive: true);

        // Assign the FontSettings instance to the document.
        doc.FontSettings = fontSettings;

        // Save the document to PDF to trigger layout and font resolution.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException($"Failed to create the PDF file at '{pdfPath}'.");
        }

        // Clean up (optional): delete the temporary custom fonts folder.
        // Directory.Delete(customFontsFolder, recursive: true);
    }
}
