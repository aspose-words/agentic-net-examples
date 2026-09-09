using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field using the typed API.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "CODE39";
        barcodeField.BarcodeValue = "12345"; // Initial data string.
        barcodeField.AddStartStopChar = true;
        builder.Writeln();

        // Ensure the field result is up‑to‑date.
        doc.UpdateFields();

        // Replace the barcode data string using Find/Replace.
        FindReplaceOptions replaceOptions = new FindReplaceOptions();
        doc.Range.Replace("12345", "ABCDE", replaceOptions);

        // Update fields again so the DISPLAYBARCODE reflects the new value.
        doc.UpdateFields();

        // Save the document.
        doc.Save("UpdatedBarcode.docx");
    }
}
