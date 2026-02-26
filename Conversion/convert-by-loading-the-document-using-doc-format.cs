// Load a Microsoft Word 97‑2003 document (.doc) using Aspose.Words
using System;
using Aspose.Words;
using Aspose.Words.Loading;

class LoadDocExample
{
    static void Main()
    {
        // Path to the source .doc file
        string sourceFile = @"C:\Docs\Sample.doc";

        // Create LoadOptions and explicitly set the format to DOC.
        // This ensures the loader treats the file as a legacy Word document.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Doc, "", "");

        // Load the document using the constructor that accepts a filename and LoadOptions.
        Document doc = new Document(sourceFile, loadOptions);

        // At this point the document is loaded and can be manipulated.
        // Example: output the first paragraph's text to the console.
        Console.WriteLine(doc.FirstSection.Body.FirstParagraph.GetText().Trim());

        // Optionally, save the document to another format (e.g., DOCX).
        // The Save method follows the standard Aspose.Words lifecycle rules.
        string outputFile = @"C:\Docs\SampleConverted.docx";
        doc.Save(outputFile, SaveFormat.Docx);
    }
}
