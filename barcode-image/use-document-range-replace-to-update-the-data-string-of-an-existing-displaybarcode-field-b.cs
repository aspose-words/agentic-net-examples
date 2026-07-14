using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field with an initial barcode value.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "QR";
        barcodeField.BarcodeValue = "ABC123";
        barcodeField.BackgroundColor = "0xF8BD69";
        barcodeField.ForegroundColor = "0xB5413B";
        barcodeField.ErrorCorrectionLevel = "3";
        barcodeField.ScalingFactor = "250";
        barcodeField.SymbolHeight = "1000";
        barcodeField.SymbolRotation = "0";

        builder.Writeln();

        // Replace the old barcode data string with a new one using Range.Replace.
        FindReplaceOptions replaceOptions = new FindReplaceOptions();
        doc.Range.Replace("ABC123", "XYZ789", replaceOptions);

        // Update fields so the DISPLAYBARCODE reflects the new value.
        doc.UpdateFields();

        // Save the updated document.
        doc.Save("UpdatedBarcode.docx");
    }
}
