using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class RenderDocWithLinuxFonts
{
    public static void Main()
    {
        // Path to a Linux fonts folder (common location). Adjust if necessary.
        string linuxFontsFolder = "/usr/share/fonts/truetype";

        // Configure FontSettings to use the Linux fonts folder.
        FontSettings fontSettings = new FontSettings();
        if (Directory.Exists(linuxFontsFolder))
        {
            // The second parameter indicates that subfolders should be searched recursively.
            fontSettings.SetFontsFolder(linuxFontsFolder, true);
        }

        // Create a simple DOC document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Use a font that is typically available in Linux installations.
        builder.Font.Name = "DejaVu Sans";
        builder.Writeln("This document is rendered on a headless Linux server using custom FontSettings.");

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        string outputPdfPath = "RenderedDocument.pdf";
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPdfPath))
            throw new FileNotFoundException("The PDF file was not created.", outputPdfPath);

        // Simple validation: check that the PDF contains an embedded font marker.
        // This does not guarantee full font embedding but confirms that the PDF was generated.
        string pdfContent = File.ReadAllText(outputPdfPath);
        if (!pdfContent.Contains("/FontFile") && !pdfContent.Contains("/FontFile2") && !pdfContent.Contains("/FontFile3"))
            Console.WriteLine("Warning: No embedded font markers were found in the PDF.");

        // Indicate successful completion.
        Console.WriteLine($"Document rendered successfully to '{outputPdfPath}'.");
    }
}
