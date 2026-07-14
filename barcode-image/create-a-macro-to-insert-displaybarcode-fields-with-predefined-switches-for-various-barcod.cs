using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    // Inserts DISPLAYBARCODE fields with predefined switches for several barcode types.
    private static void InsertBarcodeFields(DocumentBuilder builder)
    {
        // 1. QR code with custom colors and scaling.
        var qrField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        qrField.BarcodeType = "QR";
        qrField.BarcodeValue = "ABC123";
        qrField.BackgroundColor = "0xF8BD69";
        qrField.ForegroundColor = "0xB5413B";
        qrField.ErrorCorrectionLevel = "3";
        qrField.ScalingFactor = "250";
        qrField.SymbolHeight = "1000";
        qrField.SymbolRotation = "0";
        builder.Writeln();

        // 2. EAN13 barcode with displayed digits.
        var ean13Field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        ean13Field.BarcodeType = "EAN13";
        ean13Field.BarcodeValue = "501234567890";
        ean13Field.DisplayText = true;
        ean13Field.PosCodeStyle = "CASE";
        ean13Field.FixCheckDigit = true;
        builder.Writeln();

        // 3. CODE39 barcode with start/stop characters.
        var code39Field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        code39Field.BarcodeType = "CODE39";
        code39Field.BarcodeValue = "12345ABCDE";
        code39Field.AddStartStopChar = true;
        builder.Writeln();

        // 4. ITF14 barcode with a case code style.
        var itf14Field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        itf14Field.BarcodeType = "ITF14";
        itf14Field.BarcodeValue = "09312345678907";
        itf14Field.CaseCodeStyle = "STD";
        builder.Writeln();
    }

    public static void Main()
    {
        // Create a new empty document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert the predefined DISPLAYBARCODE fields.
        InsertBarcodeFields(builder);

        // Update fields to ensure the field codes are generated.
        doc.UpdateFields();

        // Save the document to the local file system.
        doc.Save("Barcodes.docx");
    }
}
