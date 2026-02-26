using System;
using Aspose.Words;

class ConvertToDocExample
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"MyDir\Document.docx";

        // Path where the converted DOC file will be saved.
        string outputPath = @"ArtifactsDir\Document.ConvertToDoc.doc";

        // Load the existing DOCX document.
        Document doc = new Document(inputPath);

        // Save the document in the legacy DOC format.
        doc.Save(outputPath, SaveFormat.Doc);
    }
}
