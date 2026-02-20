using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Create LoadOptions and explicitly set the format to DOCX.
        LoadOptions loadOptions = new LoadOptions
        {
            LoadFormat = LoadFormat.Docx
        };

        // Open the document using the specified load options.
        Document doc = new Document(sourcePath, loadOptions);

        // The document is now loaded and can be processed further.
        Console.WriteLine("Document loaded successfully. Page count: " + doc.PageCount);
    }
}
