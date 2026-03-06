using System;
using Aspose.Words;

class ConvertToDoc
{
    static void Main()
    {
        // Path to the source document (any supported format, e.g., .docx)
        string sourcePath = "input.docx";

        // Open the source document. The constructor automatically detects the format.
        Document doc = new Document(sourcePath);

        // Save the opened document in the legacy DOC format.
        string outputPath = "output.doc";
        doc.Save(outputPath, SaveFormat.Doc);
    }
}
