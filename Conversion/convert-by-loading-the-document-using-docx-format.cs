using System;
using Aspose.Words;
using Aspose.Words.Loading;

class DocxLoadAndConvert
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\source.docx";

        // Path where the converted document will be saved.
        string outputPath = @"C:\Docs\converted.pdf";

        // LoadOptions specifying that the document to be loaded is a DOCX file.
        // The second and third parameters (password, baseUri) are left empty because they are not needed.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Docx, "", "");

        // Load the DOCX document using the specified LoadOptions.
        Document doc = new Document(inputPath, loadOptions);

        // Save the loaded document as PDF (conversion).
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
