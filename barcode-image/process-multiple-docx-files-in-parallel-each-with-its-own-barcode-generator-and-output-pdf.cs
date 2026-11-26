using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Generate a barcode image using Aspose.BarCode and return it as a PNG stream.
        MemoryStream stream = new MemoryStream();
        // Use Code128 as a default symbology; parameters may contain other info if needed.
        using (BarcodeGenerator generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Optional: set colors if required.
            generator.Parameters.Barcode.BarColor = Color.Black;
            generator.Parameters.BackColor = Color.White;
            generator.Save(stream, BarCodeImageFormat.Png);
        }
        stream.Position = 0;
        return stream;
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Delegate to the primary implementation.
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare sample DOCX files.
        List<string> docxFiles = new List<string> { "doc1.docx", "doc2.docx", "doc3.docx" };
        int counter = 1;
        foreach (string filePath in docxFiles)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            // Insert a barcode field with unique value per document.
            FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            field.BarcodeType = "CODE128";
            field.BarcodeValue = $"Value{counter}";
            field.ForegroundColor = "0x000000";
            field.BackgroundColor = "0xFFFFFF";
            doc.UpdateFields();
            doc.Save(filePath);
            counter++;
        }

        // Process the DOCX files in parallel, converting each to PDF with rendered barcodes.
        Parallel.ForEach(docxFiles, docxPath =>
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);
            // Register the custom barcode generator for rendered output.
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();
            // Ensure fields are updated so barcodes are rendered.
            doc.UpdateFields();
            // Save as PDF with the same base name.
            string pdfPath = Path.ChangeExtension(docxPath, ".pdf");
            doc.Save(pdfPath);
        });
    }
}
