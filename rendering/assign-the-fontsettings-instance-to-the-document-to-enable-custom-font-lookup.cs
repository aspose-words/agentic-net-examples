using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for output and custom fonts.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        string fontsDir = Path.Combine(Directory.GetCurrentDirectory(), "MyFonts");

        // Ensure the directories exist.
        Directory.CreateDirectory(outputDir);
        Directory.CreateDirectory(fontsDir);

        // Create a new blank document.
        Document doc = new Document();

        // Add some text using a font that may not be installed on the system.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "NonExistentFont";
        builder.Writeln("This text uses a custom font lookup via FontSettings.");

        // Create a FontSettings instance and point it to the custom fonts folder.
        FontSettings fontSettings = new FontSettings();
        // The folder may be empty; Aspose.Words will also fall back to system fonts.
        fontSettings.SetFontsFolder(fontsDir, recursive: false);

        // Assign the FontSettings instance to the document.
        doc.FontSettings = fontSettings;

        // Save the document to PDF to trigger layout and font resolution.
        string pdfPath = Path.Combine(outputDir, "CustomFontLookup.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF output file.");

        // Optionally, indicate success (no interactive prompts required).
        Console.WriteLine("Document saved successfully to: " + pdfPath);
    }
}
