using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define a folder to hold custom TrueType fonts.
        string fontsFolder = Path.Combine(Directory.GetCurrentDirectory(), "fonts");
        Directory.CreateDirectory(fontsFolder);

        // Create a dummy TrueType font file in the folder.
        // In a real scenario, copy actual .ttf files here.
        string dummyFontPath = Path.Combine(fontsFolder, "DummyFont.ttf");
        if (!File.Exists(dummyFontPath))
        {
            // Write a minimal placeholder byte array.
            File.WriteAllBytes(dummyFontPath, new byte[] { 0x00, 0x01, 0x00, 0x00 });
        }

        // Configure FontSettings to use the custom fonts folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(fontsFolder, false);

        // Create a simple document that uses the custom font.
        Document doc = new Document();
        doc.FontSettings = fontSettings;
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "DummyFont";
        builder.Writeln("This text is rendered with a custom TrueType font.");

        // Render the document to PDF.
        string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "output.pdf");
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPdf))
            throw new FileNotFoundException("Failed to create the PDF output.", outputPdf);

        // Simple inspection: check for embedded font markers in the PDF content.
        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(outputPdf));
        bool hasEmbeddedFont = pdfContent.Contains("/FontFile") ||
                               pdfContent.Contains("/FontFile2") ||
                               pdfContent.Contains("/FontFile3");

        // Output the verification result.
        Console.WriteLine($"PDF generated at: {outputPdf}");
        Console.WriteLine($"Embedded font marker found: {hasEmbeddedFont}");
    }
}
