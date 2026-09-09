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

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files with heading structures.
        for (int i = 1; i <= 2; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            // First chapter heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {i} - Introduction");

            // Some normal text.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("This is some introductory content for the chapter.");

            // Second heading within the same chapter.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln($"Section {i}.1 - Details");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Detailed information goes here.");

            // Save the document as PDF.
            string pdfPath = Path.Combine(inputFolder, $"Sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Batch convert each PDF to EPUB, preserving chapter structure.
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        foreach (string pdfFile in pdfFiles)
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

            // Validate that the EPUB was created.
            if (!File.Exists(epubPath))
                throw new InvalidOperationException($"Failed to create EPUB file: {epubPath}");
        }

        // All conversions completed successfully.
    }
}
