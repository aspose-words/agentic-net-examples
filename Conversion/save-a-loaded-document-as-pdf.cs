using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing Word document from the file system.
        // The constructor automatically detects the format based on the file extension.
        Document doc = new Document("input.docx");

        // Save the loaded document as a PDF.
        // The Save method determines the output format from the ".pdf" extension.
        doc.Save("output.pdf");
    }
}
