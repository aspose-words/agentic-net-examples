using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source file (any supported format, will be loaded as DOC)
        string sourceFilePath = @"C:\Docs\sourceFile.doc";

        // Specify load options to force DOC format detection
        LoadOptions loadOptions = new LoadOptions
        {
            LoadFormat = LoadFormat.Doc
        };

        // Create a Document object from the source file using the specified load options
        Document doc = new Document(sourceFilePath, loadOptions);

        // The document is now loaded and can be used for further processing
        Console.WriteLine($"Document loaded. Page count: {doc.PageCount}");
    }
}
