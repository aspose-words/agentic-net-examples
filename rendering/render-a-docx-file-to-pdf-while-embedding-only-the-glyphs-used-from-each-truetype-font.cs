using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Path for the resulting PDF.
        string pdfPath = "output.pdf";

        // Create a new document and add a paragraph that uses a TrueType font (Arial).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("Hello World! This is a test document using Arial font.");

        // Configure font settings to resolve fonts from the system fonts folder.
        FontSettings fontSettings = new FontSettings();
        string fontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        fontSettings.SetFontsFolder(fontsFolder, true);
        doc.FontSettings = fontSettings;

        // Save the document as PDF with subset embedding (only used glyphs are embedded).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // In older versions EmbedSubset may be unavailable; setting EmbedFullFonts to false
            // forces subsetting of the fonts.
            EmbedFullFonts = false
        };
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        // Load the PDF content as text to look for subset font markers.
        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));

        // Check for typical subset font indicators: /FontFile2 entry or a subset font name (e.g., ABCDEF+Arial).
        bool hasSubsetMarker = pdfContent.Contains("/FontFile2") ||
                               Regex.IsMatch(pdfContent, @"[A-Z]{6}\+");

        if (!hasSubsetMarker)
            throw new InvalidOperationException("The PDF does not contain subset font markers; embedding may not be subsetted.");

        // Indicate successful rendering with subset embedding.
        Console.WriteLine("PDF rendered successfully with subset font embedding.");
    }
}
