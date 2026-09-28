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
        // Define temporary file paths.
        string tempDir = Path.GetTempPath();
        string docxPath = Path.Combine(tempDir, "Sample.docx");
        string pdfPath = Path.Combine(tempDir, "Sample.pdf");

        // 1. Create a sample DOCX that uses a TrueType font (e.g., Arial).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("Hello World! This is a test document using the Arial TrueType font.");
        doc.Save(docxPath, SaveFormat.Docx);

        // 2. Configure FontSettings to locate system fonts.
        FontSettings fontSettings = new FontSettings();
        string fontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        fontSettings.SetFontsFolder(fontsFolder, false);
        doc.FontSettings = fontSettings;

        // 3. Render the document to PDF with font subsetting (embed only used glyphs).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // When false, only the glyphs used in the document are embedded.
            EmbedFullFonts = false
        };
        doc.Save(pdfPath, pdfOptions);

        // 4. Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        // 5. Inspect the PDF content for subset font markers.
        // Look for patterns like "/FontFile2" or a subset prefix (e.g., "ABCDEF+ArialMT").
        string pdfText = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));

        bool hasFontFile = pdfText.Contains("/FontFile2") || pdfText.Contains("/FontFile3");
        bool hasSubsetPrefix = Regex.IsMatch(pdfText, @"\b[A-Z]{6}\+");

        if (!hasFontFile || !hasSubsetPrefix)
            throw new InvalidOperationException("The PDF does not contain expected subset font markers.");

        // 6. Indicate successful rendering and subsetting.
        Console.WriteLine("PDF rendered successfully with font subsetting.");
        Console.WriteLine($"DOCX path: {docxPath}");
        Console.WriteLine($"PDF path: {pdfPath}");
    }
}
