using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class BarcodeToPdf
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field that will render a QR code.
        // The field will generate the barcode image when fields are updated.
        FieldMergeBarcode barcodeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        barcodeField.BarcodeType = "QR";                 // QR code type
        barcodeField.BarcodeValue = "ABC123";            // Data to encode
        barcodeField.BackgroundColor = "0xF8BD69";       // Custom background color
        barcodeField.ForegroundColor = "0xB5413B";       // Custom foreground color
        barcodeField.ErrorCorrectionLevel = "3";         // QR error correction level
        barcodeField.ScalingFactor = "250";              // Scale the symbol
        barcodeField.SymbolHeight = "1000";              // Height in TWIPS
        barcodeField.SymbolRotation = "0";               // No rotation

        // Force the field to be evaluated so the barcode image is generated.
        doc.UpdateFields();

        // Save the document as PDF. The barcode image will be rendered in the PDF.
        doc.Save("BarcodeOutput.pdf", SaveFormat.Pdf);
    }
}
