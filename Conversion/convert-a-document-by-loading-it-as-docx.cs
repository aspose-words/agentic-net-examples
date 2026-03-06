using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source document (any supported format, e.g., PDF, RTF, etc.)
        string sourcePath = "source.pdf";

        // Path where the converted DOCX will be saved
        string targetPath = "converted.docx";

        // Load the source document. The constructor automatically detects the file format.
        Document doc = new Document(sourcePath);

        // Save the loaded document as DOCX.
        doc.Save(targetPath, SaveFormat.Docx);
    }
}
