using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define directories for fonts and output.
        string baseDir = Directory.GetCurrentDirectory();
        string fontsDir = Path.Combine(baseDir, "fonts");
        string outputDir = Path.Combine(baseDir, "output");

        // Ensure the directories exist.
        Directory.CreateDirectory(fontsDir);
        Directory.CreateDirectory(outputDir);

        // Create a dummy TrueType font file in the fonts directory.
        // In a real scenario, copy actual .ttf files here.
        string dummyFontPath = Path.Combine(fontsDir, "DummyFont.ttf");
        if (!File.Exists(dummyFontPath))
        {
            // Write a minimal TrueType file header (just for demonstration).
            byte[] dummyFontData = new byte[] { 0x00, 0x01, 0x00, 0x00, 0x00, 0x0C, 0x00, 0x80 };
            File.WriteAllBytes(dummyFontPath, dummyFontData);
        }

        // Create a new document and add some text using the dummy font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "DummyFont";
        builder.Writeln("This text is rendered with the DummyFont.");

        // Configure FontSettings to use the custom fonts folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(fontsDir, recursive: true);
        doc.FontSettings = fontSettings;

        // Save the document to PDF.
        string pdfPath = Path.Combine(outputDir, "RenderedDocument.pdf");
        doc.Save(pdfPath);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF output file.");

        // Optionally, you could inspect the PDF for embedded font markers here.
        // For this example, we simply confirm the file exists and exit.
    }
}
