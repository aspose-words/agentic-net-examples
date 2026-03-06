using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputFile = @"C:\Docs\Sample.docx";

        // Path to the destination PDF file.
        string outputFile = @"C:\Docs\Sample.pdf";

        // Load the DOCX document from the file system.
        Document doc = new Document(inputFile);

        // Save the document as PDF. The format is inferred from the .pdf extension.
        doc.Save(outputFile);
    }
}
