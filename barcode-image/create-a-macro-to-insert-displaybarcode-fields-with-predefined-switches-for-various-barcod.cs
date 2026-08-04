using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Drawing; // Required by Aspose.Words for barcode field properties

namespace BarcodeMacroExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert DISPLAYBARCODE fields with predefined switches.
            InsertDisplayBarcodeFields(builder);

            // Update all fields to ensure the field codes are generated.
            doc.UpdateFields();

            // Save the document to the current directory.
            doc.Save("DisplayBarcodeFields.docx");
        }

        private static void InsertDisplayBarcodeFields(DocumentBuilder builder)
        {
            // 1. QR code with custom colors and scaling.
            FieldDisplayBarcode qrField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            qrField.BarcodeType = "QR";
            qrField.BarcodeValue = "ABC123";
            qrField.BackgroundColor = "0xF8BD69";
            qrField.ForegroundColor = "0xB5413B";
            qrField.ErrorCorrectionLevel = "3";
            qrField.ScalingFactor = "250";
            qrField.SymbolHeight = "1000";
            qrField.SymbolRotation = "0";
            builder.Writeln();

            // 2. EAN13 barcode with displayed text and point‑of‑sale style.
            FieldDisplayBarcode ean13Field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            ean13Field.BarcodeType = "EAN13";
            ean13Field.BarcodeValue = "501234567890";
            ean13Field.DisplayText = true;
            ean13Field.PosCodeStyle = "CASE";
            ean13Field.FixCheckDigit = true;
            builder.Writeln();

            // 3. CODE39 barcode with start/stop characters.
            FieldDisplayBarcode code39Field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            code39Field.BarcodeType = "CODE39";
            code39Field.BarcodeValue = "12345ABCDE";
            code39Field.AddStartStopChar = true;
            builder.Writeln();

            // 4. ITF14 barcode with a case code style.
            FieldDisplayBarcode itf14Field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            itf14Field.BarcodeType = "ITF14";
            itf14Field.BarcodeValue = "09312345678907";
            itf14Field.CaseCodeStyle = "STD";
            builder.Writeln();
        }
    }
}
