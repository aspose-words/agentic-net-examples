using System;
using Aspose.Words;

class ConvertToDoc
{
    static void Main()
    {
        // Path to the source document (any supported format, e.g., DOCX).
        string sourcePath = @"C:\Docs\Document.docx";

        // Path where the converted DOC file will be saved.
        string outputPath = @"C:\Docs\Document.Converted.doc";

        // Load the source document. The constructor automatically detects the format.
        Document doc = new Document(sourcePath);

        // Save the document in the legacy Microsoft Word 97‑2007 DOC format.
        doc.Save(outputPath, SaveFormat.Doc);
    }
}
