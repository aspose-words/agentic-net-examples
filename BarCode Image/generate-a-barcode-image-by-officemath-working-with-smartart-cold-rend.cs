using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using System.Drawing;

class BarcodeSmartArtExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Insert a custom QR barcode using MERGEBARCODE field.
        // -------------------------------------------------
        // Insert the MERGEBARCODE field and cast it to FieldMergeBarcode to set its properties.
        FieldMergeBarcode barcodeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        barcodeField.BarcodeType = "QR";               // Set barcode type to QR.
        barcodeField.BarcodeValue = "MyQRCodeValue";   // The value that will be encoded.
        // Optional: customize appearance.
        barcodeField.BackgroundColor = "0xF8BD69";
        barcodeField.ForegroundColor = "0xB5413B";
        barcodeField.ErrorCorrectionLevel = "3";
        barcodeField.ScalingFactor = "250";
        barcodeField.SymbolHeight = "1000";

        // Move to a new line after the barcode.
        builder.Writeln();

        // -------------------------------------------------
        // Insert a placeholder image (demonstrates InsertImage usage).
        // -------------------------------------------------
        // Replace the path with a valid image file on your system.
        builder.InsertImage("SampleImage.png", 200, 200);
        builder.Writeln();

        // -------------------------------------------------
        // Add a text watermark to the document.
        // -------------------------------------------------
        doc.Watermark.SetText("CONFIDENTIAL");

        // -------------------------------------------------
        // Insert SmartArt (placeholder – actual SmartArt insertion requires
        // a specific API not listed in the provided rules). Here we insert a shape
        // that could later be replaced with SmartArt.
        // -------------------------------------------------
        Shape smartArtShape = builder.InsertShape(ShapeType.Rectangle, 300, 200);
        smartArtShape.WrapType = WrapType.Inline;
        // If the shape were a SmartArt object, we could trigger cold rendering like this:
        // smartArtShape.UpdateSmartArtDrawing();

        // -------------------------------------------------
        // Save the document to a DOCX file.
        // -------------------------------------------------
        doc.Save("BarcodeSmartArt.docx", SaveFormat.Docx);
    }
}
