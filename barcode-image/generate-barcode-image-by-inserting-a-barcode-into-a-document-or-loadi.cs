using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeExample
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a new document and insert a DISPLAYBARCODE field.
            // -----------------------------------------------------------------
            Document doc = new Document();                     // Create a blank document.
            DocumentBuilder builder = new DocumentBuilder(doc); // Initialize the builder.

            // Insert a DISPLAYBARCODE field. The field will display a QR code.
            FieldDisplayBarcode displayField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            displayField.BarcodeType = "QR";                    // Set barcode type.
            displayField.BarcodeValue = "ABC123";               // Set the data to encode.
            displayField.BackgroundColor = "0xF8BD69";          // Optional: set background color.
            displayField.ForegroundColor = "0xB5413B";          // Optional: set foreground color.
            displayField.ErrorCorrectionLevel = "3";           // Optional: set QR error correction level.
            displayField.ScalingFactor = "250";                // Optional: set scaling factor.
            displayField.SymbolHeight = "1000";                // Optional: set symbol height (twips).
            displayField.SymbolRotation = "0";                 // Optional: set rotation.

            // Add a paragraph break after the field for readability.
            builder.Writeln();

            // Save the newly created document with the barcode.
            doc.Save("BarcodeCreated.docx");

            // -----------------------------------------------------------------
            // 2. Load an existing DOCX document and add a BARCODE field.
            // -----------------------------------------------------------------
            // Assume there is an existing file named "Template.docx" in the same folder.
            Document existingDoc = new Document("Template.docx");
            DocumentBuilder existingBuilder = new DocumentBuilder(existingDoc);

            // Move the cursor to the end of the document (or any desired location).
            existingBuilder.MoveToDocumentEnd();

            // Insert a BARCODE field for a US ZIP code.
            FieldBarcode barcodeField = (FieldBarcode)existingBuilder.InsertField(FieldType.FieldBarcode, true);
            barcodeField.PostalAddress = "96801";   // The ZIP code to encode.
            barcodeField.IsUSPostalAddress = true; // Indicate that this is a US postal address.
            barcodeField.FacingIdentificationMark = "C"; // Optional: set FIM character.

            // Add a paragraph break after the field.
            existingBuilder.Writeln();

            // Save the modified document.
            existingDoc.Save("BarcodeAddedToExisting.docx");
        }
    }
}
