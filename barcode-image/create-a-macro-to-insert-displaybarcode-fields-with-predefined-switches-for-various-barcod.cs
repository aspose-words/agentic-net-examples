using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

namespace BarcodeFieldMacro
{
    public class Program
    {
        // Inserts DISPLAYBARCODE fields with predefined switches for several barcode types.
        private static void InsertDisplayBarcodeFields(DocumentBuilder builder)
        {
            // QR code with custom colors and scaling.
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

            // EAN13 barcode with displayed text and POS code style.
            FieldDisplayBarcode ean13Field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            ean13Field.BarcodeType = "EAN13";
            ean13Field.BarcodeValue = "501234567890";
            ean13Field.DisplayText = true;
            ean13Field.PosCodeStyle = "CASE";
            ean13Field.FixCheckDigit = true;
            builder.Writeln();

            // CODE39 barcode with start/stop characters.
            FieldDisplayBarcode code39Field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            code39Field.BarcodeType = "CODE39";
            code39Field.BarcodeValue = "12345ABCDE";
            code39Field.AddStartStopChar = true;
            builder.Writeln();

            // ITF14 barcode with case code style.
            FieldDisplayBarcode itf14Field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            itf14Field.BarcodeType = "ITF14";
            itf14Field.BarcodeValue = "09312345678907";
            itf14Field.CaseCodeStyle = "STD";
            builder.Writeln();
        }

        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the predefined DISPLAYBARCODE fields.
            InsertDisplayBarcodeFields(builder);

            // Update all fields so that their results are calculated.
            doc.UpdateFields();

            // Save the document to the local file system.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DisplayBarcodeFields.docx");
            doc.Save(outputPath, SaveFormat.Docx);

            // Indicate completion.
            Console.WriteLine("Document created: " + outputPath);
        }
    }
}
