using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a custom folder for TrueType fonts.
        string customFontFolder = Path.Combine(Directory.GetCurrentDirectory(), "CustomFonts");
        Directory.CreateDirectory(customFontFolder);

        // Attempt to copy a known system TrueType font into the custom folder (if it exists).
        // This step ensures the folder actually contains a font for rendering.
        try
        {
            string systemFontsPath = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
            string sourceFontPath = Path.Combine(systemFontsPath, "arial.ttf");
            if (File.Exists(sourceFontPath))
            {
                string destFontPath = Path.Combine(customFontFolder, "arial.ttf");
                File.Copy(sourceFontPath, destFontPath, overwrite: true);
            }
        }
        catch
        {
            // If copying fails (e.g., on non‑Windows platforms), continue without a font file.
            // Aspose.Words will fall back to its default font handling.
        }

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // Use the font we attempted to place in the custom folder.
        builder.Writeln("This text is rendered using a custom font folder.");

        // Configure FontSettings to use the custom font folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(customFontFolder, recursive: true);
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RenderedOutput.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the rendered PDF file.");

        // Simple validation: check the PDF content for embedded font markers.
        // This looks for typical PDF font embedding entries such as "/FontFile" or subset font names.
        string pdfContent = File.ReadAllText(outputPath);
        bool containsEmbeddedFont = pdfContent.Contains("/FontFile") || pdfContent.Contains("/FontFile2") ||
                                    pdfContent.Contains("/FontFile3") || System.Text.RegularExpressions.Regex.IsMatch(pdfContent, @"[A-Z]{6}\+");

        if (!containsEmbeddedFont)
            throw new InvalidOperationException("The rendered PDF does not contain embedded TrueType font markers.");

        // Indicate successful execution.
        Console.WriteLine("Document rendered successfully with custom font folder.");
    }
}
