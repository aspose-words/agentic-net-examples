using System;
using Aspose.Words;
using Aspose.Words.Fields;

class GenerateBarcodeWithField
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field that will render the barcode directly.
        // The second argument (true) tells the builder to add the field result.
        FieldDisplayBarcode displayBarcode = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Configure the barcode parameters.
        displayBarcode.BarcodeType = "QR";               // Type of barcode (QR code in this example).
        displayBarcode.BarcodeValue = "ABC123";          // Data to encode.
        displayBarcode.DisplayText = true;               // Show the encoded text below the barcode.
        displayBarcode.ForegroundColor = "0x000000";   // Black bars (hex string).
        displayBarcode.BackgroundColor = "0xFFFFFF";   // White background (hex string).
        displayBarcode.ErrorCorrectionLevel = "2";      // QR error correction level (0‑3) as string.

        // Optionally add a line break after the barcode.
        builder.Writeln();

        // Save the document in DOCX format.
        doc.Save("BarcodeField.docx");
    }
}
