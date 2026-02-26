using Aspose.Words;
using System;

class Program
{
    static void Main()
    {
        // Path to the source document (any supported format)
        string sourcePath = "input.docx";

        // Path where the DOC format file will be saved
        string outputPath = "output.doc";

        // Load the source document; format is detected automatically
        Document doc = new Document(sourcePath);

        // Save the loaded document in Microsoft Word 97‑2007 DOC format
        doc.Save(outputPath, SaveFormat.Doc);
    }
}
