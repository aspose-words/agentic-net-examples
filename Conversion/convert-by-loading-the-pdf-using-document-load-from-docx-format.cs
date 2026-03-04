using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class PdfToDocxConverter
    {
        static void Main()
        {
            // Path to the source PDF file.
            string pdfFilePath = @"C:\Input\sample.pdf";

            // Path where the resulting DOCX file will be saved.
            string docxFilePath = @"C:\Output\sample.docx";

            // Load the PDF document using PdfLoadOptions.
            PdfLoadOptions loadOptions = new PdfLoadOptions();
            Document document = new Document(pdfFilePath, loadOptions);

            // Save the loaded document as DOCX.
            document.Save(docxFilePath, SaveFormat.Docx);
        }
    }
}
