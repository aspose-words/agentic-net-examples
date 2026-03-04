using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDoc
{
    static void Main()
    {
        // Path to the source document (any supported format, e.g., DOCX).
        string inputFile = @"C:\Docs\Input.docx";

        // Path where the converted DOC file will be saved.
        string outputFile = @"C:\Docs\Output.doc";

        // Load the source document.
        Document doc = new Document(inputFile);

        // Option 1: Save directly using the SaveFormat enumeration.
        doc.Save(outputFile, SaveFormat.Doc);

        // Option 2: Use DocSaveOptions for more control (uncomment if needed).
        // DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        // doc.Save(outputFile, saveOptions);
    }
}
