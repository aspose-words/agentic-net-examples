using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a BARCODE field that displays a US ZIP code as a barcode.
        // The field will be updated automatically because we pass 'true' to InsertField.
        FieldBarcode zipBarcode = (FieldBarcode)builder.InsertField(FieldType.FieldBarcode, true);
        zipBarcode.PostalAddress = "12345";          // ZIP code to encode.
        zipBarcode.IsUSPostalAddress = true;        // Specify that this is a US postal address.
        zipBarcode.FacingIdentificationMark = "C";  // Optional FIM character.

        // Add a paragraph break after the first barcode.
        builder.Writeln();

        // Create a bookmark that contains another ZIP code.
        builder.StartBookmark("ZipBookmark");
        builder.Writeln("67890");
        builder.EndBookmark("ZipBookmark");

        // Insert a second BARCODE field that references the bookmark.
        FieldBarcode bookmarkBarcode = (FieldBarcode)builder.InsertField(FieldType.FieldBarcode, true);
        bookmarkBarcode.PostalAddress = "ZipBookmark"; // Name of the bookmark.
        bookmarkBarcode.IsBookmark = true;            // Indicate that PostalAddress is a bookmark name.
        bookmarkBarcode.IsUSPostalAddress = true;
        bookmarkBarcode.FacingIdentificationMark = "A";

        // Save the document as a DOCX file.
        doc.Save("BarCode.docx");
    }
}
