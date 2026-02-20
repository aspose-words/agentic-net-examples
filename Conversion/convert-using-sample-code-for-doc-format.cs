using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing document (any supported format).
        // The Document constructor automatically detects the format.
        Document doc = new Document("InputDocument.docx");

        // Create save options for the DOC format.
        // DocSaveOptions can be used to customize how the document is saved as .doc.
        DocSaveOptions saveOptions = new DocSaveOptions();

        // Save the document in the legacy DOC format.
        doc.Save("OutputDocument.doc", saveOptions);
    }
}
