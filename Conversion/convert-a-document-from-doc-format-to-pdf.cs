using System;
using Aspose.Words;

class DocToPdfConverter
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the PDF will be saved. The .pdf extension tells Aspose.Words to save in PDF format.
        string outputPath = @"C:\Docs\ConvertedDocument.pdf";

        // Load the existing DOC document.
        Document doc = new Document(inputPath);

        // Save the document as PDF. The format is inferred from the file extension.
        doc.Save(outputPath);
    }
}
