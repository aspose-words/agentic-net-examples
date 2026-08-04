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
        // Define output directory and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the generated PDF.
        string pdfPath = Path.Combine(outputDir, "SampleSubset.pdf");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a TrueType font that is not a core PDF font (e.g., Calibri).
        // The font will be located in the system fonts folder.
        string systemFontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        FontSettings.DefaultInstance.SetFontsFolder(systemFontsFolder, true);

        builder.Font.Name = "Calibri";
        builder.Writeln("This document uses the Calibri font.");
        builder.Writeln("Only the glyphs required for this text should be embedded.");

        // Configure PDF save options to embed only used glyphs (subsetting).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // When false (default), fonts are subsetted before embedding.
            EmbedFullFonts = false
        };

        // Save the document as PDF.
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        // Load the PDF bytes and search for markers that indicate a subsetted TrueType font.
        byte[] pdfBytes = File.ReadAllBytes(pdfPath);
        string pdfContent = Encoding.ASCII.GetString(pdfBytes);

        // Look for the typical TrueType font embedding marker "/FontFile2".
        bool hasFontFile2 = Regex.IsMatch(pdfContent, @"/FontFile2", RegexOptions.IgnoreCase);
        // Look for a subset font name pattern: six uppercase letters followed by '+'.
        bool hasSubsetName = Regex.IsMatch(pdfContent, @"[A-Z]{6}\+");

        if (!hasFontFile2 || !hasSubsetName)
            throw new InvalidOperationException("The PDF does not contain expected subsetted TrueType font markers.");

        Console.WriteLine("PDF generated successfully with subsetted TrueType fonts:");
        Console.WriteLine(pdfPath);
    }
}
