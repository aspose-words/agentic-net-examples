using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Folder where the resulting document will be saved.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Insert a DISPLAYBARCODE field and configure it.
        // -------------------------------------------------
        // The field will display a QR code with custom colors and scaling.
        FieldDisplayBarcode displayField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        displayField.BarcodeType = "QR";
        displayField.BarcodeValue = "ABC123";
        displayField.BackgroundColor = "0xF8BD69";
        displayField.ForegroundColor = "0xB5413B";
        displayField.ErrorCorrectionLevel = "3";
        displayField.ScalingFactor = "250";
        displayField.SymbolHeight = "1000";
        displayField.SymbolRotation = "0";

        // Add a line break after the field for readability.
        builder.Writeln();

        // -------------------------------------------------
        // Insert a MERGEBARCODE field for mail merge.
        // -------------------------------------------------
        // This field will be replaced by a DISPLAYBARCODE field for each data row.
        FieldMergeBarcode mergeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        mergeField.BarcodeType = "CODE39";
        mergeField.BarcodeValue = "MyCODE39Barcode";
        mergeField.AddStartStopChar = true; // Show start/stop characters.

        // Add a line break after the merge field.
        builder.Writeln();

        // -------------------------------------------------
        // Prepare mail‑merge data.
        // -------------------------------------------------
        DataTable table = new DataTable("Barcodes");
        table.Columns.Add("MyCODE39Barcode");
        table.Rows.Add(new[] { "12345ABCDE" });
        table.Rows.Add(new[] { "67890FGHIJ" });

        // Execute mail merge – each row creates a DISPLAYBARCODE field.
        doc.MailMerge.Execute(table);

        // -------------------------------------------------
        // Save the document in DOCX format.
        // -------------------------------------------------
        string outputPath = Path.Combine(artifactsDir, "BarcodeFields.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
