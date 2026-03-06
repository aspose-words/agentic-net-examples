using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input PDF file path
        string inputPdfPath = "input.pdf";

        // Output DOCX file path
        string outputDocxPath = "output.docx";

        // Create load options for PDF (default settings)
        PdfLoadOptions loadOptions = new PdfLoadOptions();

        // Load the PDF document into an Aspose.Words Document object
        Document doc = new Document(inputPdfPath, loadOptions);

        // Save the loaded document in DOCX format
        doc.Save(outputDocxPath, SaveFormat.Docx);
    }
}
