using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the source DOCX and the rendered PDF.
        string docPath = Path.Combine(outputDir, "Sample.docx");
        string pdfPath = Path.Combine(outputDir, "Sample.pdf");

        // Linux font folder (common location for TrueType fonts).
        string linuxFontsFolder = "/usr/share/fonts/truetype";

        // Create a simple document with a font that may be present on Linux.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "DejaVu Sans";
        builder.Writeln("This is a sample document rendered with custom font settings on Linux.");

        // Save the source DOCX (optional, just to have the original file).
        doc.Save(docPath);

        // Configure FontSettings to use the Linux fonts folder.
        FontSettings fontSettings = new FontSettings();
        // Scan the folder recursively to include subfolders.
        fontSettings.SetFontsFolder(linuxFontsFolder, true);
        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new Exception("PDF rendering failed: output file not found.");
        }

        // Inform the user where the PDF was saved.
        Console.WriteLine($"Rendered PDF saved to: {pdfPath}");
    }
}
