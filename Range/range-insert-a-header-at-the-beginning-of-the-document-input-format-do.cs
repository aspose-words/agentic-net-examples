using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertHeaderExample
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Create a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert the desired header text.
        builder.Writeln("My Document Header");

        // Save the modified document.
        doc.Save("OutputDocument.docx");
    }
}
