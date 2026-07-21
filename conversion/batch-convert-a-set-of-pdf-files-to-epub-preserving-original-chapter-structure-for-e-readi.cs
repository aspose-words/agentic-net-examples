using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input PDFs and output EPUBs.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchConversionDemo");
        string inputDir = Path.Combine(baseDir, "InputPdfs");
        string outputDir = Path.Combine(baseDir, "OutputEpubs");

        // Ensure clean environment.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample PDF files with heading structure.
        const int sampleCount = 3;
        for (int i = 1; i <= sampleCount; i++)
        {
            // Build a simple document with headings to represent chapters.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Chapter title (Heading 1).
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {i}: Sample Title");

            // Some body text (Normal style).
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            // Sub‑section title (Heading 2).
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln($"Section {i}.1");

            // More body text.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

            // Save the document as PDF.
            string pdfPath = Path.Combine(inputDir, $"SampleDocument{i}.pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Batch convert each PDF to EPUB while preserving chapter structure.
        string[] pdfFiles = Directory.GetFiles(inputDir, "*.pdf");
        foreach (string pdfFile in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Configure EPUB save options to split at heading paragraphs.
            HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub)
            {
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                ExportDocumentProperties = true
            };

            // Determine output EPUB path.
            string epubFileName = Path.GetFileNameWithoutExtension(pdfFile) + ".epub";
            string epubPath = Path.Combine(outputDir, epubFileName);

            // Save as EPUB.
            pdfDoc.Save(epubPath, epubOptions);

            // Validate that the EPUB was created.
            if (!File.Exists(epubPath))
                throw new InvalidOperationException($"EPUB file was not created: {epubPath}");
            if (new FileInfo(epubPath).Length == 0)
                throw new InvalidOperationException($"EPUB file is empty: {epubPath}");
        }

        // All conversions completed successfully.
        Console.WriteLine("Batch conversion completed. EPUB files are located at:");
        Console.WriteLine(outputDir);
    }
}
