using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Folder where output files will be written.
    private const string ArtifactsDir = "Artifacts";

    public static void Main()
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory(ArtifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX document that uses a TrueType font.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a common TrueType font that is available on most systems.
        builder.Font.Name = "Arial";
        builder.Writeln("Hello world! This text will be rendered with full font embedding.");

        // -----------------------------------------------------------------
        // 2. Configure PDF save options to embed the full font (no subsetting).
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // When true the whole font file is embedded, subsetting is disabled.
            EmbedFullFonts = true
        };

        // Path of the resulting PDF file.
        string pdfPath = Path.Combine(ArtifactsDir, "Document.FullFontEmbedding.pdf");

        // Save the document as PDF using the configured options.
        doc.Save(pdfPath, pdfOptions);

        Console.WriteLine($"PDF saved to: {pdfPath}");

        // -----------------------------------------------------------------
        // 3. Verify that the PDF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // -----------------------------------------------------------------
        // 4. List any extracted font files (none are expected for PDF rendering).
        // -----------------------------------------------------------------
        string[] extractedFonts = Directory.GetFiles(ArtifactsDir, "*.ttf");
        if (extractedFonts.Length > 0)
        {
            Console.WriteLine("Extracted font files:");
            foreach (string fontFile in extractedFonts)
                Console.WriteLine($" - {Path.GetFileName(fontFile)} ({new FileInfo(fontFile).Length} bytes)");
        }
        else
        {
            Console.WriteLine("No font files were extracted (PDF embedding does not produce separate .ttf files).");
        }
    }
}
