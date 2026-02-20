using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load options specific to PDF files
        var loadOptions = new PdfLoadOptions();

        // Load the PDF document using the specified options
        var document = new Document("input.pdf", loadOptions);

        // Save the loaded document in another format (e.g., DOCX)
        document.Save("output.docx", SaveFormat.Docx);
    }
}
