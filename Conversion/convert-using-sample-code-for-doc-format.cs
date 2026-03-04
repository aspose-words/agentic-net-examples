using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("MyDir/Document.docx");

        // Prepare save options for the legacy DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

        // Save the document as a .doc file.
        doc.Save("ArtifactsDir/Document.ConvertToDoc.doc", saveOptions);
    }
}
