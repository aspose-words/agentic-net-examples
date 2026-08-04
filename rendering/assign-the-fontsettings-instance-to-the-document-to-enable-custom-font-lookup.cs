using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define an output directory for generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font name that is unlikely to be present on the system.
        // This will demonstrate that the custom FontSettings are consulted during layout.
        builder.Font.Name = "NonExistentFont";
        builder.Writeln("This paragraph uses a custom font lookup.");

        // Create a FontSettings instance and assign it to the document.
        // No additional font sources are added; the instance can later be configured as needed.
        FontSettings customFontSettings = new FontSettings();
        doc.FontSettings = customFontSettings;

        // Save the document to PDF to trigger layout and font resolution.
        string pdfPath = Path.Combine(outputDir, "CustomFontLookup.pdf");
        doc.Save(pdfPath);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF output file.");

        // Optionally, inform that the process completed successfully.
        Console.WriteLine($"Document saved successfully to: {pdfPath}");
    }
}
