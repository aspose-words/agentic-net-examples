using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = "input.docx";

        // Create LoadOptions that explicitly specify the DOCX format.
        // The constructor (LoadFormat, string, string) allows setting the format,
        // password (empty here), and base URI (empty here).
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Docx, "", "");

        // Open the document using the specified LoadOptions.
        Document doc = new Document(sourcePath, loadOptions);

        // Example conversion: save the opened document as PDF.
        doc.Save("output.pdf", SaveFormat.Pdf);
    }
}
