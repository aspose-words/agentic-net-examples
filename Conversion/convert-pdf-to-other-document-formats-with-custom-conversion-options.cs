using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Paths to the source PDF and the desired output files.
        string pdfPath = @"C:\Docs\Input.pdf";
        string docxPath = @"C:\Docs\Output.docx";
        string rtfPath = @"C:\Docs\Output.rtf";

        // Load the PDF with custom load options.
        // Example: do not skip images and read only the first page.
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            SkipPdfImages = false,
            PageIndex = 0,
            PageCount = 1
        };
        Document pdfDocument = new Document(pdfPath, loadOptions);

        // Save the document as DOCX with a password using OoxmlSaveOptions.
        OoxmlSaveOptions docxSaveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Password = "SecretPassword"
        };
        pdfDocument.Save(docxPath, docxSaveOptions);

        // Save the same document as RTF without additional options.
        pdfDocument.Save(rtfPath, SaveFormat.Rtf);
    }
}
