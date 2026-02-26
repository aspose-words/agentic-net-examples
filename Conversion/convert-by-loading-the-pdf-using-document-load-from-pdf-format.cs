using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = "input.pdf";

        // Create PDF load options (default settings).
        PdfLoadOptions loadOptions = new PdfLoadOptions();

        // Load the PDF into an Aspose.Words Document object.
        Document doc = new Document(pdfPath, loadOptions);

        // Example usage: write the extracted text to the console.
        Console.WriteLine(doc.GetText());

        // Optional: save the document in another format, e.g., DOCX.
        // doc.Save("output.docx");
    }
}
