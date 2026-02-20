using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = "source.docx";

        // Create LoadOptions and explicitly set the format to DOCX.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFormat = LoadFormat.Docx;

        // Load the document using the specified options.
        Document doc = new Document(sourcePath, loadOptions);

        // The document is now loaded and can be processed further.
        // Example: output the document text to the console.
        Console.WriteLine(doc.GetText());
    }
}
