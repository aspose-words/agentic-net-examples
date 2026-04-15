using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the rendered PDF.
        string pdfPath = Path.Combine(outputDir, "CustomFontSettings.pdf");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text using a common font.
        builder.Font.Name = "Arial";
        builder.Writeln("Sample text rendered with custom FontSettings.");

        // Create a FontSettings instance and assign it to the document.
        FontSettings customFontSettings = new FontSettings();

        // (Optional) Set a custom fonts folder. Here we use an empty folder for demonstration.
        string customFontsFolder = Path.Combine(outputDir, "Fonts");
        Directory.CreateDirectory(customFontsFolder);
        customFontSettings.SetFontsFolder(customFontsFolder, recursive: false);

        // Assign the FontSettings to the document.
        doc.FontSettings = customFontSettings;

        // Render the document to PDF.
        doc.Save(pdfPath);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Indicate successful completion.
        Console.WriteLine($"PDF successfully saved to: {pdfPath}");
    }
}
