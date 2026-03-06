using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("InputDocument.docx");

        // -------------------------------------------------
        // Create a simple mail‑merge template.
        // (If the document already contains merge fields this step can be omitted.
        //  Here we add a sample MERGEFIELD for demonstration purposes.)
        // -------------------------------------------------
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();
        builder.InsertField("MERGEFIELD Name", "«Name»");
        builder.Writeln();

        // Save the document as a PNG image (renders the first page).
        doc.Save("OutputImage.png", SaveFormat.Png);
    }
}
