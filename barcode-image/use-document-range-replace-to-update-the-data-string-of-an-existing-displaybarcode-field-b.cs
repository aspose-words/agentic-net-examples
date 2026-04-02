using System;
using System.IO;
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

        // Insert a DISPLAYBARCODE field using the typed API.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "QR";
        barcodeField.BarcodeValue = "ABC123"; // Initial data string.
        barcodeField.BackgroundColor = "0xF8BD69";
        barcodeField.ForegroundColor = "0xB5413B";
        barcodeField.ErrorCorrectionLevel = "3";
        barcodeField.ScalingFactor = "250";
        barcodeField.SymbolHeight = "1000";
        barcodeField.SymbolRotation = "0";

        builder.Writeln(); // End the paragraph containing the field.

        // Replace the barcode data string using Document.Range.Replace.
        // This updates the field code text from "ABC123" to "XYZ789".
        FindReplaceOptions replaceOptions = new FindReplaceOptions();
        doc.Range.Replace("ABC123", "XYZ789", replaceOptions);

        // Recalculate fields so the DISPLAYBARCODE reflects the new value.
        doc.UpdateFields();

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "UpdatedBarcode.docx");
        doc.Save(outputPath);
    }
}
