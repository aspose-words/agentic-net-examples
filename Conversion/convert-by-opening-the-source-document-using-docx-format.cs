using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Input\SourceDocument.docx";

        // Optional: explicitly specify that the file is a DOCX when loading.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Docx, "", "");

        // Open the document using the specified load options.
        Document doc = new Document(sourcePath, loadOptions);

        // Example conversion: save the opened document as PDF.
        string outputPath = @"C:\Output\ConvertedDocument.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
