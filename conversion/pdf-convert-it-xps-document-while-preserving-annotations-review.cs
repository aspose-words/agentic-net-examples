using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToXpsConverter
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfFilePath = @"C:\Docs\SourceDocument.pdf";

        // Path where the resulting XPS file will be saved.
        string xpsFilePath = @"C:\Docs\ConvertedDocument.xps";

        // Load the PDF document into an Aspose.Words Document object.
        // The constructor automatically detects the format from the file extension.
        Document pdfDocument = new Document(pdfFilePath);

        // Create XpsSaveOptions to control the conversion.
        // No special settings are required to preserve annotations; they are kept by default.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Save the loaded document as XPS using the specified options.
        pdfDocument.Save(xpsFilePath, xpsOptions);
    }
}
