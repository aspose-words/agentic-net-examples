using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Save the document as a PDF file.
        doc.Save("Output.pdf", SaveFormat.Pdf);
    }
}
