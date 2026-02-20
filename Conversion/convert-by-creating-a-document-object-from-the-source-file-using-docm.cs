using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source DOCM file.
        string sourceFile = @"C:\Docs\SourceDocument.docm";

        // Create LoadOptions and set the format to DOCM.
        LoadOptions loadOptions = new LoadOptions
        {
            LoadFormat = LoadFormat.Docm
        };

        // Load the document using the options.
        Document doc = new Document(sourceFile, loadOptions);

        // The Document object is now ready for further processing.
        Console.WriteLine($"Document loaded. Page count: {doc.PageCount}");
    }
}
