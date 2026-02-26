using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing Word document from the file system.
        Document doc = new Document("input.docx");

        // Save the loaded document as PDF. The format is inferred from the .pdf extension.
        doc.Save("output.pdf");
    }
}
