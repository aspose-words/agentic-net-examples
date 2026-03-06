using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // Delete all characters within the document's main range.
        doc.Range.Delete();

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
