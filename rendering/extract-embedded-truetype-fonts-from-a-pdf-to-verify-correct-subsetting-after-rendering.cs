using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposePdfFontExtract");
        Directory.CreateDirectory(workFolder);

        // Path for the generated PDF.
        string pdfPath = Path.Combine(workFolder, "sample.pdf");

        // -----------------------------------------------------------------
        // 1. Create a simple document that uses a TrueType font (Arial).
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This is a test paragraph using the Arial TrueType font.");

        // Assign default FontSettings (no special embedding configuration needed;
        // Aspose.Words embeds subset fonts by default when saving to PDF).
        doc.FontSettings = new FontSettings();

        // -----------------------------------------------------------------
        // 2. Render the document to PDF.
        // -----------------------------------------------------------------
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        // -----------------------------------------------------------------
        // 3. Inspect the PDF content for embedded font markers and subset naming.
        // -----------------------------------------------------------------
        byte[] pdfBytes = File.ReadAllBytes(pdfPath);
        string pdfContent = Encoding.ASCII.GetString(pdfBytes);

        // Look for embedded font markers.
        bool hasFontFileMarker = pdfContent.Contains("/FontFile") ||
                                 pdfContent.Contains("/FontFile2") ||
                                 pdfContent.Contains("/FontFile3");

        // Look for subset font name pattern (e.g., ABCDEF+ArialMT).
        bool hasSubsetPattern = Regex.IsMatch(pdfContent, @"[A-Z]{6}\+");

        // -----------------------------------------------------------------
        // 4. Validate embedding and subsetting.
        // -----------------------------------------------------------------
        if (hasFontFileMarker && hasSubsetPattern)
        {
            Console.WriteLine("Success: PDF contains embedded TrueType font with subsetting.");
        }
        else
        {
            throw new InvalidOperationException(
                "Failed to verify embedded TrueType font or subsetting in the PDF.");
        }
    }
}
