using Aspose.Words;
using System;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Save the document as PDF. The format is inferred from the file extension.
        doc.Save("Output.pdf");
    }
}
