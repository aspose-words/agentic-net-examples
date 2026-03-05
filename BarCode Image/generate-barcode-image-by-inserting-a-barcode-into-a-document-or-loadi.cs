using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field that will generate a QR code.
        // The second argument (true) tells the builder to update the field immediately.
        FieldMergeBarcode barcodeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);

        // Set the barcode type and the value to encode.
        barcodeField.BarcodeType = "QR";               // string, not int
        barcodeField.BarcodeValue = "ABC123";          // string, not int

        // Optional visual customizations – all properties are strings.
        barcodeField.BackgroundColor = "0xF8BD69";     // string representation of color
        barcodeField.ForegroundColor = "0xB5413B";     // string representation of color
        barcodeField.ErrorCorrectionLevel = "3";      // string, even though it represents a number
        barcodeField.ScalingFactor = "250";           // string percentage value
        barcodeField.SymbolHeight = "1000";           // string height in twips
        barcodeField.SymbolRotation = "0";            // string rotation angle

        // Update the field to render the barcode image.
        barcodeField.Update();

        // Save the document in DOCX format.
        doc.Save("BarcodeDocument.docx");
    }
}
