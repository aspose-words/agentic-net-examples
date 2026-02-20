using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string sourcePath = @"C:\Docs\SourceDocument.doc";

        // Specify load options with the DOC format explicitly.
        LoadOptions loadOptions = new LoadOptions
        {
            LoadFormat = LoadFormat.Doc
        };

        // Open the document using the load options.
        Document doc = new Document(sourcePath, loadOptions);

        // (Optional) Perform any processing on the document here.
        // For demonstration, we will just output the document text to the console.
        Console.WriteLine(doc.GetText());

        // Save the document to a new file (e.g., DOCX) if needed.
        string outputPath = @"C:\Docs\ConvertedDocument.docx";
        doc.Save(outputPath);
    }
}
