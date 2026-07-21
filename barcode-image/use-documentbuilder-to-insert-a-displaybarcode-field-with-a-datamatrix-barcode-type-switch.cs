using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeFieldExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a DISPLAYBARCODE field using the typed API.
            Aspose.Words.Fields.FieldDisplayBarcode barcodeField =
                (Aspose.Words.Fields.FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

            // Set the barcode type to DataMatrix and provide a sample value.
            barcodeField.BarcodeType = "DataMatrix";
            barcodeField.BarcodeValue = "1234567890";

            // Update fields to ensure the field result is generated.
            doc.UpdateFields();

            // Prepare output directory.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // Save the document.
            string outputPath = Path.Combine(artifactsDir, "DisplayBarcodeDataMatrix.docx");
            doc.Save(outputPath);
        }
    }
}
