using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = "input.pdf";

        // Create PDF load options (customize as needed).
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            // Example option: do not skip images while loading.
            // SkipPdfImages = false
        };

        // Load the PDF document using the constructor that accepts load options.
        Document doc = new Document(pdfPath, loadOptions);

        // The document is now loaded and can be processed further.
        // For example, save it as a DOCX file:
        // doc.Save("output.docx");
    }
}
