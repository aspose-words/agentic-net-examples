// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Add a text watermark to the document.
        // -----------------------------------------------------------------
        doc.Watermark.SetText("CONFIDENTIAL");

        // -----------------------------------------------------------------
        // Insert a custom QR barcode using the MERGEBARCODE field.
        // -----------------------------------------------------------------
        // Insert the MERGEBARCODE field and cast it to the strongly‑typed class.
        FieldMergeBarcode barcode = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        barcode.BarcodeType = "QR";                     // QR code type.
        barcode.BarcodeValue = "MyQRCode";              // Value to encode.
        barcode.BackgroundColor = "0xF8BD69";           // Light orange background.
        barcode.ForegroundColor = "0xB5413B";           // Dark red foreground.
        barcode.ErrorCorrectionLevel = "3";             // Highest error correction.
        barcode.ScalingFactor = "250";                  // 250 % scaling.
        barcode.SymbolHeight = "1000";                  // Height in TWIPS.

        builder.Writeln(); // Move to a new line after the barcode.

        // -----------------------------------------------------------------
        // Insert an OfficeMath equation using an EQ field.
        // -----------------------------------------------------------------
        // Example equation: (π r) = θ φ
        builder.InsertField(@"EQ \o(\ac(\up5(\f(π)),\up5(\f(r))),\ac(\up5(\f(θ)),\up5(\f(φ))))", true);
        builder.Writeln();

        // -----------------------------------------------------------------
        // Insert a SmartArt shape and force cold rendering.
        // -----------------------------------------------------------------
        // The ShapeType.SmartArt enum value is used to create a SmartArt container.
        Shape smartArt = builder.InsertShape(ShapeType.SmartArt, 400, 200);
        // Update the SmartArt drawing using the cold rendering engine.
        smartArt.UpdateSmartArtDrawing();

        // -----------------------------------------------------------------
        // Save the document to a DOCX file.
        // -----------------------------------------------------------------
        doc.Save("BarcodeSmartArt.docx");
    }
}
