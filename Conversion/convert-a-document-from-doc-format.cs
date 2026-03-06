using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input DOC file path
        string inputFile = @"C:\Docs\Input.doc";

        // Output file path – converting to DOCX format
        string outputFile = @"C:\Docs\Output.docx";

        // Load the existing DOC document (format is auto‑detected)
        Document doc = new Document(inputFile);

        // Save the document in the desired format (DOCX in this case)
        doc.Save(outputFile, SaveFormat.Docx);
    }
}
