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
        // Temporary file paths
        string docPath = "sample.docx";
        string pdfPath = "output.pdf";

        // Configure FontSettings to use system fonts folder
        FontSettings fontSettings = new FontSettings();
        string fontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        fontSettings.SetFontsFolder(fontsFolder, false);

        // Create a simple DOCX that uses a TrueType font (Arial)
        Document doc = new Document();
        doc.FontSettings = fontSettings;
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This is a test document using the Arial TrueType font.");

        // Save the source DOCX
        doc.Save(docPath);

        // Prepare PDF save options to embed full fonts (no subsetting)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // EmbedFullFonts forces full embedding of TrueType fonts.
            EmbedFullFonts = true
            // FontEmbeddingMode is omitted because the required enum value is not available in this SDK version.
        };

        // Render the DOCX to PDF
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        // Load PDF content as text for inspection
        byte[] pdfBytes = File.ReadAllBytes(pdfPath);
        string pdfContent = Encoding.ASCII.GetString(pdfBytes);

        // Check for embedded font markers (e.g., /FontFile, /FontFile2, /FontFile3)
        bool hasEmbeddedFontMarker = pdfContent.Contains("/FontFile") ||
                                     pdfContent.Contains("/FontFile2") ||
                                     pdfContent.Contains("/FontFile3");

        // Check for subset font naming pattern (six uppercase letters followed by '+')
        bool hasSubsetFontName = Regex.IsMatch(pdfContent, @"[A-Z]{6}\+");

        // Validate that fonts are fully embedded and not subsetted
        if (!hasEmbeddedFontMarker)
            throw new Exception("The PDF does not contain embedded font markers.");

        if (hasSubsetFontName)
            throw new Exception("The PDF contains subset font names, indicating subsetting was applied.");

        // If we reach this point, verification succeeded
        Console.WriteLine("PDF rendered with full font embedding and without subsetting.");
    }
}
