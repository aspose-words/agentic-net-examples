using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert DISPLAYBARCODE fields with predefined switches.
        InsertBarcodeFields(builder);

        // Update all fields to ensure the field codes are generated.
        doc.UpdateFields();

        // Save the document to the local file system.
        doc.Save("Barcodes.docx");
    }

    private static void InsertBarcodeFields(DocumentBuilder builder)
    {
        // 1. QR code with custom colors and scaling.
        Aspose.Words.Fields.FieldDisplayBarcode qrField = (Aspose.Words.Fields.FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
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
        Aspose.Words.Fields.FieldDisplayBarcode ean13Field = (Aspose.Words.Fields.FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        ean13Field.BarcodeType = "EAN13";
        ean13Field.BarcodeValue = "501234567890";
        ean13Field.DisplayText = true;
        ean13Field.PosCodeStyle = "CASE";
        ean13Field.FixCheckDigit = true;
        builder.Writeln();

        // 3. CODE39 barcode with start/stop characters.
        Aspose.Words.Fields.FieldDisplayBarcode code39Field = (Aspose.Words.Fields.FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        code39Field.BarcodeType = "CODE39";
        code39Field.BarcodeValue = "12345ABCDE";
        code39Field.AddStartStopChar = true;
        builder.Writeln();

        // 4. ITF14 barcode with a specified case code style.
        Aspose.Words.Fields.FieldDisplayBarcode itf14Field = (Aspose.Words.Fields.FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        itf14Field.BarcodeType = "ITF14";
        itf14Field.BarcodeValue = "09312345678907";
        itf14Field.CaseCodeStyle = "STD";
        builder.Writeln();
    }
}
