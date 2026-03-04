using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a paragraph in a DOCM file.");

        // Save the document as a macro‑enabled DOCM file.
        string docmPath = "Document.docm";
        OoxmlSaveOptions docmSaveOptions = new OoxmlSaveOptions(SaveFormat.Docm);
        doc.Save(docmPath, docmSaveOptions);

        // Load the previously saved DOCM document.
        Document loadedDoc = new Document(docmPath);

        // Save the loaded document as a legacy DOC file.
        string docPath = "Document.doc";
        loadedDoc.Save(docPath, SaveFormat.Doc);
    }
}
