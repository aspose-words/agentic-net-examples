using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfToDocxConverter
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Temp\SourceDocument.pdf";

        // Path where the converted DOCX will be saved.
        string docxPath = @"C:\Temp\ConvertedDocument.docx";

        // Load the PDF with options that help preserve the original layout.
        // PdfLoadOptions is a derived class of LoadOptions specific to PDF files.
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            // Converting metafile images (WMF/EMF) to PNG helps keep their appearance.
            ConvertMetafilesToPng = true,

            // Preserve the original page size and margins.
            // (No explicit property for this; the default behavior already retains layout.)
        };

        // Load the PDF document using the specified load options.
        Document pdfDocument = new Document(pdfPath, loadOptions);

        // Rebuild the page layout to ensure that the document's internal pagination
        // reflects the PDF's original layout before conversion.
        pdfDocument.UpdatePageLayout();

        // Save the document as DOCX. The file extension determines the format,
        // but we explicitly specify SaveFormat.Docx for clarity.
        pdfDocument.Save(docxPath, SaveFormat.Docx);
    }
}
