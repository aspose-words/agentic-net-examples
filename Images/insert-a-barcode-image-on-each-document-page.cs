using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;

class InsertBarcodeOnEachPage
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the primary header so the content appears on every page.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a DISPLAYBARCODE field. The second argument (true) updates the field result immediately.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Configure the barcode properties.
        barcodeField.BarcodeType = "QR";                     // Type of barcode (QR, CODE39, EAN13, etc.).
        barcodeField.BarcodeValue = "ABC123";                // Data to encode.
        barcodeField.BackgroundColor = "0xF8BD69";           // Optional background color.
        barcodeField.ForegroundColor = "0xB5413B";           // Optional foreground color.
        barcodeField.ErrorCorrectionLevel = "3";            // QR specific: error correction level.
        barcodeField.ScalingFactor = "250";                 // Scale the symbol.
        barcodeField.SymbolHeight = "1000";                 // Height in twips.
        barcodeField.SymbolRotation = "0";                  // Rotation.

        // Optionally, you can force an update of all fields in the document.
        doc.UpdateFields();

        // Save the document.
        doc.Save("BarcodeOnEachPage.docx");
    }
}
