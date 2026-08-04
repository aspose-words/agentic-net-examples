using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input PDFs and output EPUBs.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputEpubs");

        // Ensure clean environment.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files with heading structure.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            // Chapter heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {i}");

            // Sub‑section heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln($"Section {i}.1");
            builder.Writeln($"Section {i}.2");

            // Normal paragraph.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit.");

            string pdfPath = Path.Combine(inputFolder, $"Sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Batch convert each PDF to EPUB, preserving chapter structure.
        foreach (string pdfFile in Directory.GetFiles(inputFolder, "*.pdf"))
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Configure EPUB save options to split at heading paragraphs.
            HtmlSaveOptions epubOptions = new HtmlSaveOptions
            {
                SaveFormat = SaveFormat.Epub,
                Encoding = Encoding.UTF8,
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                ExportDocumentProperties = true
            };

            // Determine output EPUB path.
            string epubFileName = Path.GetFileNameWithoutExtension(pdfFile) + ".epub";
            string epubPath = Path.Combine(outputFolder, epubFileName);

            // Save as EPUB.
            pdfDoc.Save(epubPath, epubOptions);

            // Validate output.
            if (!File.Exists(epubPath))
                throw new InvalidOperationException($"EPUB file was not created: {epubPath}");
        }

        // Optional: confirm that all EPUB files were generated.
        int epubCount = Directory.GetFiles(outputFolder, "*.epub").Length;
        if (epubCount == 0)
            throw new InvalidOperationException("No EPUB files were generated.");

        // Program completes without waiting for user input.
    }
}
