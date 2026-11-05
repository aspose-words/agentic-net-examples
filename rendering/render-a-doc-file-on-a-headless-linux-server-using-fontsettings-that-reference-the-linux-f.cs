using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Paths for the sample document and rendered PDF.
        string sourcePath = "sample.docx";
        string outputPath = "rendered.pdf";

        // 1. Create a simple DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Use a font that is commonly available on Linux (e.g., DejaVu Sans).
        builder.Font.Name = "DejaVu Sans";
        builder.Writeln("Hello, world! This document is rendered on a headless Linux server.");
        doc.Save(sourcePath);

        // 2. Load the document with FontSettings pointing to the Linux fonts folder.
        Document loadedDoc = new Document(sourcePath);
        FontSettings fontSettings = new FontSettings();
        // Set the folder that contains system fonts on most Linux distributions.
        // The second parameter (true) enables recursive search in subfolders.
        fontSettings.SetFontsFolder("/usr/share/fonts", true);
        loadedDoc.FontSettings = fontSettings;

        // 3. Render the document to PDF.
        loadedDoc.Save(outputPath, SaveFormat.Pdf);

        // 4. Verify that the PDF was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the rendered PDF at '{outputPath}'.");
        }

        // Optional: indicate success (no interactive prompts).
        Console.WriteLine("Document rendered successfully to PDF.");
    }
}
