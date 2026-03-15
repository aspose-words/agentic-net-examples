using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsBarcodeDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a new blank document and insert a DISPLAYBARCODE field.
            // -----------------------------------------------------------------
            Document doc = new Document();                     // Create a blank document.
            DocumentBuilder builder = new DocumentBuilder(doc); // Helper to add content.

            // Insert a DISPLAYBARCODE field that will render a QR code.
            // The field is inserted with the "true" flag to update its result immediately.
            FieldDisplayBarcode displayField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            displayField.BarcodeType = "QR";                   // Set barcode type.
            displayField.BarcodeValue = "ABC123";              // Data to encode.
            displayField.BackgroundColor = "0xF8BD69";         // Optional background color.
            displayField.ForegroundColor = "0xB5413B";         // Optional foreground color.
            displayField.ErrorCorrectionLevel = "3";           // QR error correction level (0‑3).
            displayField.ScalingFactor = "250";                // Scale the symbol (percentage).
            displayField.SymbolHeight = "1000";                // Height in TWIPS (1/1440 inch).
            displayField.SymbolRotation = "0";                 // No rotation.

            // Add a paragraph break after the field for readability.
            builder.Writeln();

            // Save the document in DOCX format.
            doc.Save("DisplayBarcode.docx");

            // -----------------------------------------------------------------
            // 2. Load an existing document and insert a BARCODE field (U.S. ZIP code).
            // -----------------------------------------------------------------
            // Assume "Template.docx" exists in the same folder.
            Document loadedDoc = new Document("Template.docx");
            DocumentBuilder loadBuilder = new DocumentBuilder(loadedDoc);

            // Move the cursor to the end of the document.
            loadBuilder.MoveToDocumentEnd();

            // Insert a BARCODE field that displays a ZIP code as a postal barcode.
            FieldBarcode barcodeField = (FieldBarcode)loadBuilder.InsertField(FieldType.FieldBarcode, true);
            barcodeField.PostalAddress = "96801";   // ZIP code to encode.
            barcodeField.IsUSPostalAddress = true; // Indicate it's a U.S. postal address.
            barcodeField.FacingIdentificationMark = "C"; // Optional FIM character.

            // Save the modified document.
            loadedDoc.Save("LoadedWithBarcode.docx");
        }
    }
}
