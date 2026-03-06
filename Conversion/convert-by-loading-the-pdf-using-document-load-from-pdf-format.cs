using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the PDF file to be loaded.
        string pdfPath = "input.pdf";

        // Configure PDF load options (optional).
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            // Load images from the PDF.
            SkipPdfImages = false,
            // Start loading from the first page (0‑based index).
            PageIndex = 0,
            // Load all pages (default is MaxValue).
            PageCount = int.MaxValue
        };

        // Load the PDF document into an Aspose.Words Document object.
        Document doc = new Document(pdfPath, loadOptions);

        // Example usage: write the extracted text to the console.
        Console.WriteLine(doc.GetText());

        // Optional: save the loaded document to another format, e.g., DOCX.
        // doc.Save("output.docx");
    }
}
