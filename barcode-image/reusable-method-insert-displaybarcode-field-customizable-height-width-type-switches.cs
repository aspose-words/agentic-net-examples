using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeDemo
{
    public static class BarcodeHelper
    {
        /// <summary>
        /// Inserts a DISPLAYBARCODE field into the document using the supplied parameters.
        /// </summary>
        /// <param name="builder">DocumentBuilder positioned where the field should be inserted.</param>
        /// <param name="value">The data to be encoded in the barcode.</param>
        /// <param name="type">Barcode type (e.g., "QR", "EAN13", "CODE39", "ITF14").</param>
        /// <param name="symbolHeight">Height of the barcode symbol in TWIPS (1/1440 inch).</param>
        /// <param name="scalingFactor">Scaling factor in whole percentage points (10‑1000).</param>
        /// <param name="backgroundColor">Optional background color in hex (e.g., "0xF8BD69").</param>
        /// <param name="foregroundColor">Optional foreground color in hex (e.g., "0xB5413B").</param>
        /// <param name="errorCorrectionLevel">Optional error correction level for QR codes (0‑3).</param>
        /// <param name="symbolRotation">Optional rotation (0‑3).</param>
        /// <param name="addStartStopChar">Optional flag for CODE39/NW7 start‑stop characters.</param>
        /// <param name="displayText">Optional flag to display the barcode data as text.</param>
        /// <param name="posCodeStyle">Optional POS code style for EAN/UPC barcodes.</param>
        /// <param name="fixCheckDigit">Optional flag to fix an invalid check digit.</param>
        /// <param name="caseCodeStyle">Optional case code style for ITF14 barcodes.</param>
        /// <returns>The created FieldDisplayBarcode instance.</returns>
        public static FieldDisplayBarcode InsertDisplayBarcode(
            DocumentBuilder builder,
            string value,
            string type,
            string symbolHeight,
            string scalingFactor,
            string backgroundColor = null,
            string foregroundColor = null,
            string errorCorrectionLevel = null,
            string symbolRotation = null,
            bool? addStartStopChar = null,
            bool? displayText = null,
            string posCodeStyle = null,
            bool? fixCheckDigit = null,
            string caseCodeStyle = null)
        {
            // Insert the field; the second argument (true) tells the builder to add field separators.
            FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

            // Required properties.
            field.BarcodeValue = value;
            field.BarcodeType = type;
            field.SymbolHeight = symbolHeight;
            field.ScalingFactor = scalingFactor;

            // Optional properties – set only when a value is supplied.
            if (!string.IsNullOrEmpty(backgroundColor))
                field.BackgroundColor = backgroundColor;

            if (!string.IsNullOrEmpty(foregroundColor))
                field.ForegroundColor = foregroundColor;

            if (!string.IsNullOrEmpty(errorCorrectionLevel))
                field.ErrorCorrectionLevel = errorCorrectionLevel;

            if (!string.IsNullOrEmpty(symbolRotation))
                field.SymbolRotation = symbolRotation;

            if (addStartStopChar.HasValue)
                field.AddStartStopChar = addStartStopChar.Value;

            if (displayText.HasValue)
                field.DisplayText = displayText.Value;

            if (!string.IsNullOrEmpty(posCodeStyle))
                field.PosCodeStyle = posCodeStyle;

            if (fixCheckDigit.HasValue)
                field.FixCheckDigit = fixCheckDigit.Value;

            if (!string.IsNullOrEmpty(caseCodeStyle))
                field.CaseCodeStyle = caseCodeStyle;

            return field;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Example: Insert a QR code with custom size and colors.
            BarcodeHelper.InsertDisplayBarcode(
                builder,
                value: "ABC123",
                type: "QR",
                symbolHeight: "1000",      // 1000 TWIPS ≈ 0.69 inch
                scalingFactor: "250",      // 250%
                backgroundColor: "0xF8BD69",
                foregroundColor: "0xB5413B",
                errorCorrectionLevel: "3",
                symbolRotation: "0");

            builder.Writeln(); // Move to next line.

            // Example: Insert an EAN13 barcode with text displayed below.
            BarcodeHelper.InsertDisplayBarcode(
                builder,
                value: "501234567890",
                type: "EAN13",
                symbolHeight: "800",
                scalingFactor: "200",
                displayText: true,
                posCodeStyle: "CASE",
                fixCheckDigit: true);

            builder.Writeln();

            // Example: Insert a CODE39 barcode with start/stop characters.
            BarcodeHelper.InsertDisplayBarcode(
                builder,
                value: "12345ABCDE",
                type: "CODE39",
                symbolHeight: "900",
                scalingFactor: "150",
                addStartStopChar: true);

            builder.Writeln();

            // Example: Insert an ITF14 barcode with a case code style.
            BarcodeHelper.InsertDisplayBarcode(
                builder,
                value: "09312345678907",
                type: "ITF14",
                symbolHeight: "1100",
                scalingFactor: "300",
                caseCodeStyle: "STD");

            // Save the document.
            doc.Save("DisplayBarcodes.docx");
        }
    }
}
