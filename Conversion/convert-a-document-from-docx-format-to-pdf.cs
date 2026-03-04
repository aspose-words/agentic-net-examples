using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputFile = "input.docx";

        // Path where the PDF will be saved.
        string outputFile = "output.pdf";

        // Load the DOCX document.
        Document doc = new Document(inputFile);

        // Save the document as PDF. The format is inferred from the file extension.
        doc.Save(outputFile);
    }
}
