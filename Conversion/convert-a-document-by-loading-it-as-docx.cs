using System;
using Aspose.Words;

class DocumentConverter
{
    static void Main()
    {
        // Path to the source document (can be any supported format)
        string sourcePath = @"C:\Docs\SourceDocument.pdf";

        // Path where the converted DOCX will be saved
        string targetPath = @"C:\Docs\ConvertedDocument.docx";

        // Load the source document. The constructor automatically detects the format.
        Document doc = new Document(sourcePath);

        // Save the loaded document as DOCX. The .docx extension tells Aspose.Words the desired format.
        doc.Save(targetPath);
    }
}
