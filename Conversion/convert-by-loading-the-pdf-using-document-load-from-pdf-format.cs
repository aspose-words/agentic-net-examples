using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = "input.pdf";

        // Create PdfLoadOptions – you can customize options here if needed.
        PdfLoadOptions loadOptions = new PdfLoadOptions();

        // Load the PDF document into an Aspose.Words Document object.
        Document doc = new Document(pdfPath, loadOptions);

        // Example: save the loaded document as a DOCX file.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
