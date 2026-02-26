using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("input.docx");

        // Update all fields in the document.
        // This ensures that any "SEE ALSO" fields (or other fields) are refreshed.
        doc.UpdateFields();

        // Save the document as a PDF file.
        doc.Save("output.pdf", SaveFormat.Pdf);
    }
}
