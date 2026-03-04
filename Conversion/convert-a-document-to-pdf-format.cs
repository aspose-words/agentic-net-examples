using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (DOCX, DOC, etc.) from the file system.
        // The Document constructor automatically detects the format.
        Document doc = new Document("input.docx");

        // Save the loaded document as PDF.
        // The Save method with a file name and SaveFormat explicitly specifies the output format.
        doc.Save("output.pdf", SaveFormat.Pdf);
    }
}
