using System;
using Aspose.Words;

class DocxToPdfConverter
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputFile = @"C:\Docs\Sample.docx";

        // Path where the resulting PDF will be saved.
        string outputFile = @"C:\Docs\Sample.pdf";

        // Load the DOCX document.
        Document doc = new Document(inputFile);

        // Save the document as PDF. The format is inferred from the file extension.
        doc.Save(outputFile);
    }
}
