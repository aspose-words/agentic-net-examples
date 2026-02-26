using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source document (any supported format, e.g., .docx, .pdf, .rtf, etc.)
        string inputPath = "input.docx";

        // Path where the converted DOC file will be saved.
        string outputPath = "output.doc";

        // Load the source document. Aspose.Words automatically detects the format.
        Document doc = new Document(inputPath);

        // Create save options for the legacy DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

        // Save the document to the target DOC file.
        doc.Save(outputPath, saveOptions);
    }
}
