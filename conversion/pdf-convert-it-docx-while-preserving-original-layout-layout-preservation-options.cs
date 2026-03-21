using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfToDocxConverter
{
    static void Main()
    {
        // Use the system temporary folder for all files.
        string tempFolder = Path.GetTempPath();

        // Path to the source PDF file.
        string pdfPath = Path.Combine(tempFolder, "SourceDocument.pdf");

        // Path where the converted DOCX will be saved.
        string docxPath = Path.Combine(tempFolder, "ConvertedDocument.docx");

        // If the PDF does not exist, create a simple one to demonstrate the conversion.
        if (!File.Exists(pdfPath))
        {
            Document tempDoc = new Document();
            var builder = new DocumentBuilder(tempDoc);
            builder.Writeln("This is a sample PDF created for conversion demonstration.");
            builder.Writeln("It contains a few lines of text to illustrate layout preservation.");
            tempDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Load the PDF with options that help preserve the original layout.
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            // Converting metafile images (WMF/EMF) to PNG helps keep their appearance.
            ConvertMetafilesToPng = true
        };

        // Load the PDF document using the specified load options.
        Document pdfDocument = new Document(pdfPath, loadOptions);

        // Rebuild the page layout to ensure that the document's internal pagination
        // reflects the PDF's original layout before conversion.
        pdfDocument.UpdatePageLayout();

        // Save the document as DOCX.
        pdfDocument.Save(docxPath, SaveFormat.Docx);

        Console.WriteLine($"PDF successfully converted to DOCX:");
        Console.WriteLine($"PDF : {pdfPath}");
        Console.WriteLine($"DOCX: {docxPath}");
    }
}
