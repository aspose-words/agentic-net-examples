using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchPdfToEpubConverter
{
    public static void Main()
    {
        // Prepare input folder and create sample PDF files with chapter headings.
        string inputFolder = "InputPdfs";
        Directory.CreateDirectory(inputFolder);

        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            // Create a heading to represent a chapter.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {i}");

            // Add some content under the heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"This is the content of chapter {i}.");

            string pdfPath = Path.Combine(inputFolder, $"Sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Prepare output folder for EPUB files.
        string outputFolder = "OutputEpubs";
        Directory.CreateDirectory(outputFolder);

        // Batch convert each PDF in the input folder to EPUB.
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        foreach (string pdfFilePath in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDocument = new Document(pdfFilePath);

            // Determine output EPUB path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(pdfFilePath);
            string epubPath = Path.Combine(outputFolder, $"{fileNameWithoutExt}.epub");

            // Save as EPUB, preserving the document structure.
            pdfDocument.Save(epubPath, SaveFormat.Epub);

            // Validate that the EPUB file was created.
            if (!File.Exists(epubPath))
            {
                throw new InvalidOperationException($"EPUB file was not created: {epubPath}");
            }
        }

        // Optional: indicate successful completion (no interactive input).
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
