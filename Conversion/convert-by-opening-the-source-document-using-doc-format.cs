using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source document (any supported format, e.g., .docx)
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Path where the converted DOC format will be saved
        string targetPath = @"C:\Docs\ConvertedDocument.doc";

        // Load the source document. The constructor automatically detects the format.
        Document doc = new Document(sourcePath);

        // Save the document in Microsoft Word 97‑2007 DOC format.
        doc.Save(targetPath, SaveFormat.Doc);
    }
}
