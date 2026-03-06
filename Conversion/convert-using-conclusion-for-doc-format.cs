using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Define input and output folders (adjust paths as needed).
        string inputFolder = @"C:\Docs\Input\";
        string outputFolder = @"C:\Docs\Output\";

        // Load the source document (any supported format, e.g., DOCX).
        Document doc = new Document(inputFolder + "SourceDocument.docx");

        // Save the document in the legacy DOC format.
        // The Save method overload with (string, SaveFormat) follows the provided lifecycle rule.
        doc.Save(outputFolder + "ConvertedDocument.doc", SaveFormat.Doc);
    }
}
