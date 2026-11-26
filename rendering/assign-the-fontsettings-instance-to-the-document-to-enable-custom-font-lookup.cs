using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define paths for the output document and a custom fonts folder.
        string outputPath = "RenderedDocument.pdf";
        string fontsFolder = "CustomFonts";

        // Ensure the custom fonts folder exists (it can be empty for this example).
        Directory.CreateDirectory(fontsFolder);

        // Create a FontSettings instance and point it to the custom fonts folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(fontsFolder, false);

        // Create a new blank document.
        Document doc = new Document();

        // Assign the FontSettings to the document to enable custom font lookup.
        doc.FontSettings = fontSettings;

        // Add a simple paragraph to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document uses custom FontSettings.");

        // Save the document to PDF format.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Optionally, output a confirmation (no user interaction required).
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
