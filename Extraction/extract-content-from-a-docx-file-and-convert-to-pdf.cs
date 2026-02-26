using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file
        string sourcePath = "input.docx";

        // Path where the PDF will be saved
        string destinationPath = "output.pdf";

        // Load the DOCX document from the file system
        Document doc = new Document(sourcePath);

        // Convert and save the document as PDF
        doc.Save(destinationPath, SaveFormat.Pdf);
    }
}
