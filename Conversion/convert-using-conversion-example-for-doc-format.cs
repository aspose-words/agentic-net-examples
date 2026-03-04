using Aspose.Words;
using Aspose.Words.Saving;
using System.IO;

class ConvertToDoc
{
    static void Main()
    {
        // Input DOCX file path
        string inputPath = @"C:\Docs\Document.docx";

        // Output DOC file path
        string outputPath = @"C:\Docs\Converted\Document.doc";

        // Load the source document (DOCX)
        Document doc = new Document(inputPath);

        // Save the document in DOC format
        doc.Save(outputPath, SaveFormat.Doc);
    }
}
