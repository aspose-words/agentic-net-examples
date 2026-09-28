using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the sample files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeWordsFontExtractionSample");
        Directory.CreateDirectory(tempFolder);

        // Path for the generated PDF.
        string pdfPath = Path.Combine(tempFolder, "sample.pdf");

        // Build a simple document that uses a TrueType font (e.g., Arial).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This is a sample text to test TrueType font embedding and subsetting.");

        // Configure PDF save options to embed subset fonts (default behavior).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure fonts are embedded as subsets rather than full fonts.
            EmbedFullFonts = false
        };

        // Render the document to PDF.
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("The PDF file was not generated.", pdfPath);

        // Read the PDF content as text for inspection.
        byte[] pdfBytes = File.ReadAllBytes(pdfPath);
        string pdfContent = Encoding.ASCII.GetString(pdfBytes);

        // Look for markers that indicate embedded TrueType fonts.
        bool containsFontFileMarker = pdfContent.Contains("/FontFile") ||
                                      pdfContent.Contains("/FontFile2") ||
                                      pdfContent.Contains("/FontFile3");

        bool containsTrueTypeSubtype = pdfContent.Contains("/Subtype /TrueType");

        // Subset fonts are usually named with six uppercase letters followed by '+' (e.g., ABCDEF+ArialMT).
        bool containsSubsetFontName = Regex.IsMatch(pdfContent, @"[A-Z]{6}\+");

        // Validate that at least one of the expected markers is present.
        if (!containsFontFileMarker && !containsTrueTypeSubtype && !containsSubsetFontName)
            throw new Exception("No embedded TrueType font markers were found in the generated PDF.");

        // If we reach this point, the verification succeeded.
        Console.WriteLine("Embedded TrueType font markers detected successfully.");
    }
}
