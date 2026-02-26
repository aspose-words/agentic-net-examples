using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source document (DOC format)
        string sourcePath = @"C:\Docs\SourceDocument.doc";

        // Create LoadOptions and explicitly set the LoadFormat to DOC.
        // This forces Aspose.Words to treat the input as a legacy Word document.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Doc, "", "");

        // Load the document using the constructor that accepts a filename and LoadOptions.
        Document doc = new Document(sourcePath, loadOptions);

        // At this point the document is loaded and can be processed.
        // Example: output the first paragraph's text to the console.
        Console.WriteLine(doc.FirstSection.Body.FirstParagraph.GetText().Trim());
    }
}
