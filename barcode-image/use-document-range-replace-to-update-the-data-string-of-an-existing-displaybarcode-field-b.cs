using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field with an initial value.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "QR";
        barcodeField.BarcodeValue = "OLDVALUE";
        barcodeField.BackgroundColor = "0xF8BD69";
        barcodeField.ForegroundColor = "0xB5413B";
        barcodeField.ErrorCorrectionLevel = "3";
        barcodeField.ScalingFactor = "250";
        barcodeField.SymbolHeight = "1000";
        barcodeField.SymbolRotation = "0";

        builder.Writeln();

        // Replace the old barcode data string with a new one.
        FindReplaceOptions options = new FindReplaceOptions();
        doc.Range.Replace("OLDVALUE", "NEWVALUE", options);

        // Update fields to reflect the new barcode value.
        doc.UpdateFields();

        // Save the document.
        doc.Save("UpdatedBarCode.docx");
    }
}
