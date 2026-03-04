using System;
using Aspose.Words;
using Aspose.Words.Fields;

class GenerateBarcode
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field. This field will generate a barcode when the document is processed.
        // The second argument (true) tells the builder to update the field result immediately.
        FieldMergeBarcode barcodeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);

        // Specify the type of barcode you want to generate (e.g., QR, CODE39, EAN13, etc.).
        barcodeField.BarcodeType = "QR";

        // Set the value that will be encoded in the barcode.
        barcodeField.BarcodeValue = "ABC123";

        // Optional: customize appearance (background, foreground colors, error correction, scaling, etc.).
        // Uncomment and adjust as needed.
        // barcodeField.BackgroundColor = "0xF8BD69";
        // barcodeField.ForegroundColor = "0xB5413B";
        // barcodeField.ErrorCorrectionLevel = "3";
        // barcodeField.ScalingFactor = "250";
        // barcodeField.SymbolHeight = "1000";
        // barcodeField.SymbolRotation = "0";

        // Add a paragraph break after the field so the barcode appears on its own line.
        builder.Writeln();

        // Save the document to a DOCX file.
        doc.Save("Barcode.docx");
    }
}
