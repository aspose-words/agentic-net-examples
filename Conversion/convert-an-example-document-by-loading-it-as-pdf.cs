using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source PDF file.
        const string inputPath = "input.pdf";

        // Path where the converted PDF will be saved.
        const string outputPath = "output.pdf";

        // Load the PDF document using PdfLoadOptions (optional, can be omitted for default loading).
        LoadOptions loadOptions = new PdfLoadOptions();
        Document doc = new Document(inputPath, loadOptions);

        // Save the loaded document back to PDF format.
        doc.Save(outputPath);
    }
}
