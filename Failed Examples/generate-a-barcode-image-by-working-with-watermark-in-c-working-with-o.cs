// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

namespace AsposeWordsBarcodeWatermarkExample
{
    // Simple barcode generator that creates a placeholder image.
    // In a real scenario you would integrate a barcode library.
    public class SimpleBarcodeGenerator : IBarcodeGenerator
    {
        public Image GetBarcodeImage(BarcodeParameters parameters)
        {
            // Create a bitmap with the barcode value drawn as text.
            const int width = 300;
            const int height = 100;
            var bitmap = new Bitmap(width, height);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (var font = new Font("Arial", 24, FontStyle.Bold))
                using (var brush = new SolidBrush(Color.Black))
                {
                    var text = parameters.BarcodeValue ?? "BARCODE";
                    var size = graphics.MeasureString(text, font);
                    var point = new PointF((width - size.Width) / 2, (height - size.Height) / 2);
                    graphics.DrawString(text, font, brush, point);
                }
            }
            return bitmap;
        }

        public Image GetOldBarcodeImage(BarcodeParameters parameters)
        {
            // For simplicity, reuse the same implementation.
            return GetBarcodeImage(parameters);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Assign the custom barcode generator to the document's field options.
            doc.FieldOptions.BarcodeGenerator = new SimpleBarcodeGenerator();

            // Generate a barcode image using the generator.
            var barcodeParams = new BarcodeParameters
            {
                BarcodeType = "CODE128",
                BarcodeValue = "1234567890"
            };
            Image barcodeImage = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParams);

            // Add the generated barcode image as a watermark.
            var imageOptions = new ImageWatermarkOptions
            {
                Scale = 0.5,          // Scale the image to 50% of its original size.
                IsWashout = false    // Make the watermark fully opaque.
            };
            doc.Watermark.SetImage(barcodeImage, imageOptions);

            // Insert an OLE object (e.g., an embedded Excel workbook) into the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            // The file path should point to an existing file you want to embed.
            // Here we use a placeholder path; replace with a real file when running the code.
            string oleFilePath = @"C:\Temp\Sample.xlsx";
            if (File.Exists(oleFilePath))
            {
                // Insert the OLE object. The progId for Excel is "Excel.Sheet.12".
                builder.InsertOleObject(oleFilePath, "Excel.Sheet.12", false, "Embedded Excel");
            }

            // Insert a hyperlink to an online video.
            string videoUrl = "https://www.example.com/video.mp4";
            builder.Writeln(); // Ensure we are on a new paragraph.
            builder.InsertHyperlink("Watch Video", videoUrl, false);
            builder.Writeln(); // Add a line break after the link.

            // Save the document to a DOCX file.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "BarcodeWatermarkWithOleAndVideo.docx");
            doc.Save(outputPath, SaveFormat.Docx);
        }
    }
}
