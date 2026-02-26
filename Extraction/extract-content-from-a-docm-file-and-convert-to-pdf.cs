using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the macro-enabled Word document (DOCM) from the file system.
        Document doc = new Document("input.docm");

        // Convert and save the document as a PDF file.
        // The SaveFormat.Pdf enum specifies the PDF output format.
        doc.Save("output.pdf", SaveFormat.Pdf);
    }
}
