using System;
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeDemo
{
    // Entry point required for a console application.
    public static void Main(string[] args)
    {
        Run();
    }

    public static void Run()
    {
        // Load the base DOCX document.
        Document original = new Document("Template.docx");

        // Clone the loaded document to work on a separate copy.
        Document cloned = original.Clone();

        // Use DocumentBuilder to insert a MERGEBARCODE field at the end of the cloned document.
        DocumentBuilder builder = new DocumentBuilder(cloned);
        builder.MoveToDocumentEnd();

        // Insert the MERGEBARCODE field and configure its properties.
        FieldMergeBarcode barcodeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        barcodeField.BarcodeType = "QR";                     // QR code type.
        barcodeField.BarcodeValue = "MyQRCode";              // Data to encode.
        barcodeField.BackgroundColor = "0xF8BD69";           // Background colour.
        barcodeField.ForegroundColor = "0xB5413B";           // Foreground colour.
        barcodeField.ErrorCorrectionLevel = "3";             // QR error correction.
        barcodeField.ScalingFactor = "250";                  // Scale the symbol.
        barcodeField.SymbolHeight = "1000";                  // Height in TWIPS.
        barcodeField.SymbolRotation = "0";                   // No rotation.

        // Force field update so the barcode image is generated.
        cloned.UpdateFields();

        // Load another document that will be appended to the cloned document.
        Document toAppend = new Document("Appendix.docx");
        cloned.AppendDocument(toAppend, ImportFormatMode.KeepSourceFormatting);

        // Split the combined document into two parts:
        // Part 1 – pages 1 and 2 (zero‑based page indices).
        Document part1 = cloned.ExtractPages(0, 1);
        // Part 2 – remaining pages.
        Document part2 = cloned.ExtractPages(2, cloned.PageCount - 1);

        // Save the resulting documents.
        cloned.Save("Combined.docx");
        part1.Save("Part1.docx");
        part2.Save("Part2.docx");
    }
}
