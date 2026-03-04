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
        FieldMergeBarcode barcodeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        barcodeField.BarcodeType = "QR";                 // QR code type
        barcodeField.BarcodeValue = "ABC123";            // Data to encode
        barcodeField.BackgroundColor = "0xF8BD69";       // Background colour
        barcodeField.ForegroundColor = "0xB5413B";       // Foreground colour
        barcodeField.ErrorCorrectionLevel = "3";         // QR error correction level
        barcodeField.ScalingFactor = "250";              // Scaling factor (percentage)
        barcodeField.SymbolHeight = "1000";              // Height in twips
        barcodeField.SymbolRotation = "0";               // Rotation

        // Force the field to update so the barcode image is generated.
        barcodeField.Update();

        // Save the document as PDF. The barcode will be rendered in the PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save("BarcodeOutput.pdf", pdfOptions);
    }
}
