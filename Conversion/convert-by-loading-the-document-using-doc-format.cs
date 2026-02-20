using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Specify load options with the DOC format explicitly.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFormat = LoadFormat.Doc; // Microsoft Word 95/97-2003 document

        // Load the document from a file using the specified options.
        Document doc = new Document("InputDocument.doc", loadOptions);

        // The document is now loaded and can be processed further.
        // Example: output the document text to the console.
        Console.WriteLine(doc.GetText());
    }
}
