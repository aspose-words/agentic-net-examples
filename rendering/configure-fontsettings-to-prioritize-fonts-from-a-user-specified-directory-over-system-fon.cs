using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare paths
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string pdfPath = Path.Combine(outputDir, "RenderedDocument.pdf");
        string customFontsDir = Path.Combine(outputDir, "CustomFonts");
        Directory.CreateDirectory(customFontsDir);

        // Attempt to copy a system font (Arial) into the custom fonts folder.
        // This makes the custom folder contain a usable font for the example.
        try
        {
            string systemFontsPath = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
            string sourceFontPath = Path.Combine(systemFontsPath, "arial.ttf");
            if (File.Exists(sourceFontPath))
            {
                string destFontPath = Path.Combine(customFontsDir, "arial.ttf");
                File.Copy(sourceFontPath, destFontPath, true);
            }
        }
        catch
        {
            // If copying fails, continue; the example will still demonstrate FontSettings configuration.
        }

        // Create a simple document that uses the Arial font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This text should be rendered using the Arial font from the custom fonts folder.");

        // Configure FontSettings to prioritize the custom fonts folder.
        FontSettings fontSettings = new FontSettings();
        // Add custom folder first – it gets higher priority.
        fontSettings.SetFontsFolder(customFontsDir, true);
        // Optionally add the system fonts folder as a fallback.
        string systemFontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        if (Directory.Exists(systemFontsFolder))
            fontSettings.SetFontsFolder(systemFontsFolder, true);

        // Assign the FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        doc.Save(pdfPath);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF rendering failed; output file not found.");

        // Simple validation: check that the PDF contains a reference to the Arial font.
        // This looks for the typical subset font name pattern (e.g., "ABCDEF+Arial").
        string pdfContent = File.ReadAllText(pdfPath);
        if (!pdfContent.Contains("Arial"))
            throw new InvalidOperationException("Rendered PDF does not contain the expected font reference.");

        // Clean up (optional): comment out if you want to inspect the files after execution.
        // Directory.Delete(outputDir, true);
    }
}
