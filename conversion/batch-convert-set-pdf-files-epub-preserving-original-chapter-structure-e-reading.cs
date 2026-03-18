using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToEpubBatch
{
    static void Main()
    {
        // Folder containing source PDF files
        string inputFolder = @"C:\InputPdfs";
        // Folder where the resulting EPUB files will be placed
        string outputFolder = @"C:\OutputEpubs";

        Directory.CreateDirectory(outputFolder);

        // Process each PDF file in the input folder
        foreach (string pdfPath in Directory.GetFiles(inputFolder, "*.pdf"))
        {
            // Load the PDF document
            Document pdfDoc = new Document(pdfPath);

            // Configure EPUB save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.SaveFormat = SaveFormat.Epub;          // Target format
            saveOptions.Encoding = Encoding.UTF8;              // Use UTF‑8 encoding
            saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph; // Preserve chapter structure
            saveOptions.ExportDocumentProperties = true;      // Optional: include document properties

            // Determine output file name (same base name, .epub extension)
            string fileName = Path.GetFileNameWithoutExtension(pdfPath);
            string epubPath = Path.Combine(outputFolder, fileName + ".epub");

            // Save the document as EPUB using the configured options
            pdfDoc.Save(epubPath, saveOptions);
        }
    }
}
