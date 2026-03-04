using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source document (DOCX, DOC, etc.).
        Document doc = new Document("input.docx");

        // Save the document as PDF. The format is inferred from the .pdf extension.
        doc.Save("output.pdf");
    }
}
