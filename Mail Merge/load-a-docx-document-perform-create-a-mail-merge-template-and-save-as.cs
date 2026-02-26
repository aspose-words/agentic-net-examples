using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("input.docx");

        // Create a simple mail‑merge template by inserting merge fields.
        // This demonstrates how to prepare the document for a later mail merge.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(); // start a new paragraph
        builder.InsertField("MERGEFIELD FirstName"); // insert first name field
        builder.Write(" "); // space between fields
        builder.InsertField("MERGEFIELD LastName"); // insert last name field

        // Save the first page of the document as a PNG image.
        // The Save overload with SaveFormat renders the document to an image.
        doc.Save("output.png", SaveFormat.Png);
    }
}
