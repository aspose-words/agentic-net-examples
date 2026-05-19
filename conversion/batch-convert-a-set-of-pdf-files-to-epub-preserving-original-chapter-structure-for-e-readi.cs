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
        string inputFolder = "InputPdfs";
        string outputFolder = "OutputEpubs";

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files with heading structures.
        CreateSamplePdf(Path.Combine(inputFolder, "Sample1.pdf"));
        CreateSamplePdf(Path.Combine(inputFolder, "Sample2.pdf"));

        // Get all PDF files in the input folder.
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");

        foreach (string pdfPath in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfPath);

            // Configure EPUB save options to split at heading paragraphs (preserves chapters).
            HtmlSaveOptions epubOptions = new HtmlSaveOptions
            {
                SaveFormat = SaveFormat.Epub,
                Encoding = Encoding.UTF8,
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                ExportDocumentProperties = true
            };

            // Determine output EPUB file name.
            string epubFileName = Path.GetFileNameWithoutExtension(pdfPath) + ".epub";
            string epubPath = Path.Combine(outputFolder, epubFileName);

            // Save as EPUB.
            pdfDoc.Save(epubPath, epubOptions);

            // Validate that the EPUB file was created.
            if (!File.Exists(epubPath))
                throw new InvalidOperationException($"EPUB file was not created: {epubPath}");
        }

        // Optional: indicate successful completion (no console interaction required).
        // The program will exit automatically.
    }

    // Helper method to create a sample PDF with headings representing chapters.
    private static void CreateSamplePdf(string outputPdfPath)
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Chapter 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the introduction chapter content.");

        // Chapter 2
        builder.InsertBreak(BreakType.PageBreak);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2: Details");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Detailed information goes here.");

        // Save the document as PDF.
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException($"PDF file was not created: {outputPdfPath}");
    }
}
