using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX document from the file system.
        Document doc = new Document("InputDocument.docx");

        // Save the loaded document as PDF.
        // The Save method automatically determines the format from the file extension.
        doc.Save("OutputDocument.pdf");
    }
}
