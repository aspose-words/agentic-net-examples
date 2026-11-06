using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        MemoryStream stream = new MemoryStream();
        using (BarcodeGenerator generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Set colors if needed
            generator.Parameters.Barcode.BarColor = Color.Black;
            generator.Parameters.BackColor = Color.White;
            generator.Save(stream, BarCodeImageFormat.Png);
        }
        stream.Position = 0;
        return stream;
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare temporary directories
        string baseDir = Path.Combine(Path.GetTempPath(), "AsposeBarcodeDemo");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOCX files with a barcode field
        List<string> inputFiles = new List<string>();
        for (int i = 1; i <= 5; i++)
        {
            string docPath = Path.Combine(inputDir, $"Document{i}.docx");
            CreateSampleDocx(docPath, $"CODE128", $"DOC{i:D3}");
            inputFiles.Add(docPath);
        }

        // Process DOCX files in parallel, each with its own barcode generator, and save PDFs
        Parallel.ForEach(inputFiles, inputPath =>
        {
            // Load the DOCX
            Document doc = new Document(inputPath);

            // Register a custom barcode generator for this document
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            // Update fields to render barcodes
            doc.UpdateFields();

            // Determine output PDF path
            string fileName = Path.GetFileNameWithoutExtension(inputPath);
            string pdfPath = Path.Combine(outputDir, $"{fileName}.pdf");

            // Save as PDF
            doc.Save(pdfPath);
        });

        // Optionally, clean up temporary files (commented out to allow inspection)
        // Directory.Delete(baseDir, true);
    }

    private static void CreateSampleDocx(string filePath, string barcodeType, string barcodeValue)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a barcode field using typed insertion
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = barcodeType;          // e.g., "CODE128"
        field.BarcodeValue = barcodeValue;        // e.g., "1234567890"
        field.ForegroundColor = "0x000000";       // Black
        field.BackgroundColor = "0xFFFFFF";       // White

        // Save the DOCX
        doc.Save(filePath);
    }
}
