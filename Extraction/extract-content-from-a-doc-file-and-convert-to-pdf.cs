using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the existing DOC file. The constructor automatically detects the format.
        Document doc = new Document("input.doc");

        // Save the document as PDF. The file extension determines the output format.
        doc.Save("output.pdf");
    }
}
