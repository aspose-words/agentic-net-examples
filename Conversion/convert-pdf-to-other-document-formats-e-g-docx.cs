using System;
using Aspose.Words;

class PdfConverter
{
    // Converts a PDF file to the format indicated by the output file extension.
    public static void ConvertPdf(string pdfPath, string outputPath)
    {
        // Load the PDF document from the file system.
        Document pdfDocument = new Document(pdfPath);

        // Save the document; the format is inferred from the output file extension.
        pdfDocument.Save(outputPath);
    }

    // Example usage.
    static void Main()
    {
        // Path to the source PDF.
        string sourcePdf = @"C:\Docs\sample.pdf";

        // Path to the desired output (DOCX in this case).
        string targetDocx = @"C:\Docs\sample_converted.docx";

        // Perform the conversion.
        ConvertPdf(sourcePdf, targetDocx);
    }
}
