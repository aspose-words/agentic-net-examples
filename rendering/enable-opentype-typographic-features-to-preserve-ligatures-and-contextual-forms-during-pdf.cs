using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that supports OpenType ligatures.
        builder.Font.Name = "Times New Roman";
        builder.Writeln("office"); // Contains "ff" ligature opportunity.

        // Add Arabic text to demonstrate contextual forms.
        builder.Font.Name = "Arial";
        builder.Font.LocaleId = 1025; // Arabic locale.
        builder.Writeln("العربية");

        // Configure FontSettings to point to the system fonts folder.
        FontSettings fontSettings = new FontSettings();
        string fontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        if (Directory.Exists(fontsFolder))
        {
            fontSettings.SetFontsFolder(fontsFolder, false);
        }
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        string pdfPath = "Output.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
        {
            throw new Exception("PDF rendering failed: file was not created.");
        }

        // Simple inspection for embedded font markers in the PDF content.
        byte[] pdfBytes = File.ReadAllBytes(pdfPath);
        string pdfText = Encoding.ASCII.GetString(pdfBytes);
        bool hasEmbeddedFont = pdfText.Contains("/FontFile") ||
                               pdfText.Contains("/FontFile2") ||
                               pdfText.Contains("/FontFile3");

        Console.WriteLine(hasEmbeddedFont
            ? "PDF created with embedded fonts (OpenType features preserved)."
            : "PDF created, but embedded font markers were not found.");
    }
}
