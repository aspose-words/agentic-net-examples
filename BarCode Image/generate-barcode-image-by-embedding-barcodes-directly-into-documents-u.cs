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

        // Insert a MERGEBARCODE field that will render a QR code.
        // The second argument (true) tells the builder to update the field immediately.
        FieldMergeBarcode barcodeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);

        // Set the type of barcode to QR and provide the data to encode.
        barcodeField.BarcodeType = "QR";
        barcodeField.BarcodeValue = "https://example.com";

        // Optional visual customizations.
        // BackgroundColor and ForegroundColor expect a string with a hex color value (without the leading '#').
        barcodeField.BackgroundColor = "F8BD69";          // Light orange background.
        barcodeField.ForegroundColor = "B5413B";          // Dark red bars.
        barcodeField.ErrorCorrectionLevel = "3";         // Highest error correction.
        barcodeField.ScalingFactor = "250";              // 250 % scaling.
        barcodeField.SymbolHeight = "1000";              // Height in TWIPS (≈0.69 in).
        barcodeField.SymbolRotation = "0";               // No rotation.

        // Ensure the field result is calculated (necessary if the field was not updated on insert).
        doc.UpdateFields();

        // Save the document in DOCX format.
        doc.Save("BarcodeMerge.docx");
    }
}
