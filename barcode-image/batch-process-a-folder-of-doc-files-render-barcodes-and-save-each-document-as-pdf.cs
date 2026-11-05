using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a PNG barcode image from the given parameters.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        MemoryStream stream = new MemoryStream();
        using (BarcodeGenerator generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Set barcode and background colors.
            generator.Parameters.Barcode.BarColor = Color.Black;
            generator.Parameters.BackColor = Color.White;
            generator.Save(stream, BarCodeImageFormat.Png);
        }
        stream.Position = 0;
        return stream;
    }

    // Legacy method – delegate to the primary implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    // Folder containing DOC/DOCX files to process.
    private const string InputFolder = "Docs";

    // Output folder for generated PDFs.
    private const string OutputFolder = "PdfOutput";

    public static void Main()
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory(OutputFolder);

        // Verify the input directory exists; if not, exit gracefully.
        if (!Directory.Exists(InputFolder))
        {
            Console.WriteLine($"Input folder \"{InputFolder}\" does not exist. No files to process.");
            return;
        }

        // Collect .doc and .docx files from the input folder.
        List<string> allFiles = new List<string>();
        allFiles.AddRange(Directory.GetFiles(InputFolder, "*.doc", SearchOption.TopDirectoryOnly));
        allFiles.AddRange(Directory.GetFiles(InputFolder, "*.docx", SearchOption.TopDirectoryOnly));

        if (allFiles.Count == 0)
        {
            Console.WriteLine($"No DOC/DOCX files found in \"{InputFolder}\".");
            return;
        }

        foreach (string filePath in allFiles)
        {
            // Load the Word document.
            Document doc = new Document(filePath);

            // Register the custom barcode generator for rendered output (PDF).
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            // Insert a barcode field at the end of the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();

            FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            barcodeField.BarcodeType = "CODE128";
            // Use the file name (without extension) as the barcode value.
            barcodeField.BarcodeValue = Path.GetFileNameWithoutExtension(filePath);
            // Set colors using the required hex string format.
            barcodeField.ForegroundColor = "0x000000";
            barcodeField.BackgroundColor = "0xFFFFFF";

            // Render the barcode by updating fields.
            doc.UpdateFields();

            // Determine the output PDF path.
            string outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".pdf";
            string outputPath = Path.Combine(OutputFolder, outputFileName);

            // Save the document as PDF.
            doc.Save(outputPath, SaveFormat.Pdf);
        }

        Console.WriteLine("Batch processing completed successfully.");
    }
}
