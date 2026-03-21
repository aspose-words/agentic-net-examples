using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToEpubBatch
{
    static void Main()
    {
        // Base directory of the application
        string baseDir = AppContext.BaseDirectory;

        // Folder containing source PDF files (created if missing)
        string inputFolder = Path.Combine(baseDir, "InputPdfs");
        Directory.CreateDirectory(inputFolder);

        // Folder where the resulting EPUB files will be placed (created if missing)
        string outputFolder = Path.Combine(baseDir, "OutputEpubs");
        Directory.CreateDirectory(outputFolder);

        // Get all PDF files in the input folder
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        if (pdfFiles.Length == 0)
        {
            Console.WriteLine($"No PDF files found in '{inputFolder}'. Place PDFs there and rerun.");
            return;
        }

        // Process each PDF file
        foreach (string pdfPath in pdfFiles)
        {
            // Load the PDF document
            Document pdfDoc = new Document(pdfPath);

            // Configure EPUB save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                SaveFormat = SaveFormat.Epub,               // Target format
                Encoding = Encoding.UTF8,                   // Use UTF‑8 encoding
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph, // Preserve chapter structure
                ExportDocumentProperties = true            // Include document properties (optional)
            };

            // Determine output file name (same base name, .epub extension)
            string fileName = Path.GetFileNameWithoutExtension(pdfPath);
            string epubPath = Path.Combine(outputFolder, fileName + ".epub");

            // Save the document as EPUB using the configured options
            pdfDoc.Save(epubPath, saveOptions);
            Console.WriteLine($"Converted '{pdfPath}' → '{epubPath}'.");
        }
    }
}
