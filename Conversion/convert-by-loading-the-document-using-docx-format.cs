using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = "input.docx";

        // LoadOptions specifying that the document format is DOCX.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Docx, "", "");

        // Load the document using the specified LoadOptions.
        Document doc = new Document(sourcePath, loadOptions);

        // Example conversion: save the loaded document as PDF.
        string outputPath = "output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
