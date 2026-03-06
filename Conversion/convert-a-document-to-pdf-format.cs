using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source document (any format supported by Aspose.Words)
        string inputPath = "input.docx";

        // Desired output PDF file path; the .pdf extension tells Aspose.Words to save as PDF
        string outputPath = "output.pdf";

        // Load the document from the file system
        Document doc = new Document(inputPath);

        // Save the loaded document as PDF
        doc.Save(outputPath);
    }
}
