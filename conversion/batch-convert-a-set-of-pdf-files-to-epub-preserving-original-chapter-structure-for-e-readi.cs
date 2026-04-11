using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input PDFs and output EPUBs.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputPdfs");
        string outputDir = Path.Combine(baseDir, "OutputEpubs");

        // Ensure clean directories.
        if (Directory.Exists(inputDir))
            Directory.Delete(inputDir, true);
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample PDF files with heading structure.
        const int sampleCount = 3;
        for (int i = 1; i <= sampleCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a main heading (chapter).
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {i}");

            // Add a sub‑heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln($"Section {i}.1");

            // Add regular paragraph content.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            string pdfPath = Path.Combine(inputDir, $"Sample{i}.pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Batch convert each PDF to EPUB while preserving chapter structure.
        string[] pdfFiles = Directory.GetFiles(inputDir, "*.pdf");
        foreach (string pdfFile in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Configure EPUB save options.
            HtmlSaveOptions epubOptions = new HtmlSaveOptions
            {
                SaveFormat = SaveFormat.Epub,
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                ExportDocumentProperties = true
            };

            // Determine output EPUB path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(pdfFile);
            string epubPath = Path.Combine(outputDir, $"{fileNameWithoutExt}.epub");

            // Save as EPUB.
            pdfDoc.Save(epubPath, epubOptions);

            // Validate that the EPUB file was created.
            if (!File.Exists(epubPath) || new FileInfo(epubPath).Length == 0)
                throw new InvalidOperationException($"Failed to create EPUB for '{pdfFile}'.");
        }

        // Indicate completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
