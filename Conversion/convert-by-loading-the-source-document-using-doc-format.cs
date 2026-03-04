using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string sourcePath = @"C:\Docs\SourceDocument.doc";

        // LoadOptions with explicit LoadFormat set to DOC.
        LoadOptions loadOptions = new LoadOptions
        {
            LoadFormat = LoadFormat.Doc
        };

        // Load the document using the constructor that accepts a file name and LoadOptions.
        Document doc = new Document(sourcePath, loadOptions);

        // Example: save the loaded document as DOCX (optional, demonstrates saving).
        string outputPath = @"C:\Docs\ConvertedDocument.docx";
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
