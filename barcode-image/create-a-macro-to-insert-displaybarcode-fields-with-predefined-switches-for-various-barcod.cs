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

        // 1. QR code with custom colors, scaling and rotation.
        InsertDisplayBarcode(
            builder,
            barcodeValue: "ABC123",
            barcodeType: "QR",
            backgroundColor: "0xF8BD69",
            foregroundColor: "0xB5413B",
            errorCorrectionLevel: "3",
            scalingFactor: "250",
            symbolHeight: "1000",
            symbolRotation: "0",
            displayText: false);

        builder.Writeln();

        // 2. EAN13 barcode with digits displayed below the bars.
        InsertDisplayBarcode(
            builder,
            barcodeValue: "501234567890",
            barcodeType: "EAN13",
            displayText: true,
            posCodeStyle: "CASE",
            fixCheckDigit: true);

        builder.Writeln();

        // 3. CODE39 barcode with start/stop characters.
        InsertDisplayBarcode(
            builder,
            barcodeValue: "12345ABCDE",
            barcodeType: "CODE39",
            addStartStopChar: true);

        builder.Writeln();

        // 4. ITF14 barcode with a case code style.
        InsertDisplayBarcode(
            builder,
            barcodeValue: "09312345678907",
            barcodeType: "ITF14",
            caseCodeStyle: "STD");

        // Update all fields to ensure they render correctly.
        doc.UpdateFields();

        // Save the document.
        doc.Save("DisplayBarcode.docx");
    }

    private static void InsertDisplayBarcode(
        DocumentBuilder builder,
        string barcodeValue,
        string barcodeType,
        string backgroundColor = null,
        string foregroundColor = null,
        string errorCorrectionLevel = null,
        string scalingFactor = null,
        string symbolHeight = null,
        string symbolRotation = null,
        bool displayText = false,
        string posCodeStyle = null,
        bool fixCheckDigit = false,
        string caseCodeStyle = null,
        bool addStartStopChar = false)
    {
        // Insert a DISPLAYBARCODE field using the typed API.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Mandatory properties.
        field.BarcodeValue = barcodeValue;
        field.BarcodeType = barcodeType;

        // Optional properties – set only when a value is supplied.
        if (!string.IsNullOrEmpty(backgroundColor))
            field.BackgroundColor = backgroundColor;
        if (!string.IsNullOrEmpty(foregroundColor))
            field.ForegroundColor = foregroundColor;
        if (!string.IsNullOrEmpty(errorCorrectionLevel))
            field.ErrorCorrectionLevel = errorCorrectionLevel;
        if (!string.IsNullOrEmpty(scalingFactor))
            field.ScalingFactor = scalingFactor;
        if (!string.IsNullOrEmpty(symbolHeight))
            field.SymbolHeight = symbolHeight;
        if (!string.IsNullOrEmpty(symbolRotation))
            field.SymbolRotation = symbolRotation;

        field.DisplayText = displayText;

        if (!string.IsNullOrEmpty(posCodeStyle))
            field.PosCodeStyle = posCodeStyle;

        field.FixCheckDigit = fixCheckDigit;

        if (!string.IsNullOrEmpty(caseCodeStyle))
            field.CaseCodeStyle = caseCodeStyle;

        if (addStartStopChar)
            field.AddStartStopChar = true;

        // Force the field to recalculate its result.
        field.Update();
    }
}
