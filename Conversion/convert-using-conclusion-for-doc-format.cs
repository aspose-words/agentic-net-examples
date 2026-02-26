using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document. The format is detected automatically.
        Document doc = new Document("InputDocument.docx"); // TODO: replace with your source file path

        // Save the document in the legacy Microsoft Word 97‑2007 DOC format.
        doc.Save("ConvertedDocument.doc", SaveFormat.Doc); // TODO: replace with your desired output path
    }
}
