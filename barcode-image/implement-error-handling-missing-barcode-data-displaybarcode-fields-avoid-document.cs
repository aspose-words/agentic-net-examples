using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class DisplayBarcodeHandler
{
    static void Main()
    {
        // Use files in the current working directory to avoid hard‑coded paths.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");

        Document doc;

        // If the input file does not exist, create a simple document with an empty DISPLAYBARCODE field.
        if (File.Exists(inputPath))
        {
            doc = new Document(inputPath);
        }
        else
        {
            doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Sample document containing a DISPLAYBARCODE field with missing data:");
            // Insert a DISPLAYBARCODE field with an empty value.
            builder.InsertField("DISPLAYBARCODE \"\"");
            doc.Save(inputPath);
        }

        // Iterate through all fields in the document.
        foreach (Field field in doc.Range.Fields)
        {
            // Process only DISPLAYBARCODE fields.
            if (field.Type == FieldType.FieldDisplayBarcode)
            {
                var barcodeField = (FieldDisplayBarcode)field;

                // If the barcode value is null, empty, or whitespace, replace it with a placeholder.
                if (string.IsNullOrWhiteSpace(barcodeField.BarcodeValue))
                {
                    barcodeField.BarcodeValue = "N/A";
                }
            }
        }

        // Save the processed document.
        try
        {
            doc.Save(outputPath);
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving document: {ex.Message}");
        }
    }
}
