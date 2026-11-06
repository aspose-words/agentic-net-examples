using System;
using System.IO;
using Aspose.Words;

public class BatchPdfToEpubConverter
{
    public static void Main()
    {
        // Define input and output directories.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputEpubs");

        // Ensure directories exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files with chapter headings.
        for (int i = 1; i <= 2; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            // Chapter 1
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Sample PDF {i} - Chapter 1");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("This is the first chapter content.");

            // Chapter 2
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Sample PDF {i} - Chapter 2");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("This is the second chapter content.");

            string pdfPath = Path.Combine(inputFolder, $"Sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);

            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create sample PDF: {pdfPath}");
        }

        // Batch convert each PDF to EPUB.
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        foreach (string pdfFile in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Determine output EPUB path.
            string epubFileName = Path.ChangeExtension(Path.GetFileNameWithoutExtension(pdfFile), ".epub");
            string epubPath = Path.Combine(outputFolder, epubFileName);

            // Save as EPUB, preserving headings as chapters.
            pdfDoc.Save(epubPath, SaveFormat.Epub);

            if (!File.Exists(epubPath))
                throw new InvalidOperationException($"EPUB conversion failed for: {pdfFile}");
        }

        // Indicate completion.
        Console.WriteLine("Batch PDF to EPUB conversion completed successfully.");
    }
}
