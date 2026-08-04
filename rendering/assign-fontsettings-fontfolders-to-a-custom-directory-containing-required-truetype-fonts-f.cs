using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a base directory for all generated files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(baseDir);

        // Create a custom folder that will act as the font source.
        string customFontsDir = Path.Combine(baseDir, "CustomFonts");
        Directory.CreateDirectory(customFontsDir);

        // (Optional) If you have a TrueType font file you can copy it into the folder here.
        // For this example we simply leave the folder empty to demonstrate the configuration.

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font name that is not guaranteed to be present on the system.
        builder.Font.Name = "Alte DIN 1451 Mittelschrift";
        builder.Writeln("This text is rendered using a custom font folder.");

        // Configure FontSettings to use the custom fonts directory.
        FontSettings fontSettings = new FontSettings();
        // Scan subfolders recursively (set to true) – adjust as needed.
        fontSettings.SetFontsFolder(customFontsDir, recursive: true);
        doc.FontSettings = fontSettings;

        // Save the document to PDF to trigger rendering with the specified font settings.
        string outputPath = Path.Combine(baseDir, "RenderedDocument.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Indicate successful completion.
        Console.WriteLine("Document rendered and saved to: " + outputPath);
    }
}
