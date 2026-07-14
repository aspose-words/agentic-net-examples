using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input PDFs and output EPUBs.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputEpubs");

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files with heading structure.
        const int sampleCount = 3;
        for (int i = 1; i <= sampleCount; i++)
        {
            // Build a simple document with a heading (chapter) and some body text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {i}: Sample Title");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"This is the content of chapter {i}. It contains several sentences to simulate a real document.");

            // Save the document as PDF.
            string pdfPath = Path.Combine(inputFolder, $"Sample{i}.pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Batch convert each PDF to EPUB while preserving chapter (heading) structure.
        foreach (string pdfFile in Directory.GetFiles(inputFolder, "*.pdf"))
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Configure EPUB save options to split the output at heading paragraphs.
            HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub)
            {
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                ExportDocumentProperties = true
            };

            // Determine output EPUB file name.
            string epubFileName = Path.GetFileNameWithoutExtension(pdfFile) + ".epub";
            string epubPath = Path.Combine(outputFolder, epubFileName);

            // Save as EPUB.
            pdfDoc.Save(epubPath, epubOptions);

            // Verify that the EPUB file was created.
            if (!File.Exists(epubPath))
                throw new InvalidOperationException($"EPUB conversion failed for '{pdfFile}'.");
        }

        // All conversions completed successfully.
        Console.WriteLine("Batch PDF to EPUB conversion completed.");
    }
}
