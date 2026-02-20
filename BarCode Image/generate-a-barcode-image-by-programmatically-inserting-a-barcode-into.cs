using System;
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeInsertionExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field.
        // Syntax: DISPLAYBARCODE <BarcodeType> "<BarcodeValue>"
        // Example: DISPLAYBARCODE QR "ABC123"
        builder.InsertField(@"DISPLAYBARCODE QR ""ABC123""");

        // Insert a paragraph break between the two fields.
        builder.Writeln();

        // Insert a MERGEBARCODE field.
        // Syntax: MERGEBARCODE <BarcodeType> "<BarcodeValue>"
        // Example: MERGEBARCODE QR "ABC123"
        builder.InsertField(@"MERGEBARCODE QR ""ABC123""");

        // Update all fields in the document so that the barcode images are generated.
        doc.UpdateFields();

        // Save the document in DOCX format.
        doc.Save("Barcodes.docx");
    }
}
