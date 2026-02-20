using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX file, explicitly specifying the format.
        var loadOptions = new LoadOptions { LoadFormat = LoadFormat.Docx };
        Document doc = new Document("input.docx", loadOptions);

        // Save the loaded document as PDF using PdfSaveOptions.
        var saveOptions = new PdfSaveOptions();
        doc.Save("output.pdf", saveOptions);
    }
}
