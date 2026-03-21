using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeLoggingExample
{
    // Custom barcode generator that logs each generation request.
    public class LoggingBarcodeGenerator : IBarcodeGenerator
    {
        private readonly IBarcodeGenerator _innerGenerator;

        public LoggingBarcodeGenerator(IBarcodeGenerator innerGenerator = null)
        {
            _innerGenerator = innerGenerator;
        }

        // Generate barcode image for DISPLAYBARCODE fields.
        public Stream GetBarcodeImage(BarcodeParameters parameters)
        {
            Stream imageStream = _innerGenerator?.GetBarcodeImage(parameters) ?? CreatePlaceholderImage();

            // Ensure the stream is positioned at the beginning.
            if (imageStream.CanSeek)
                imageStream.Position = 0;

            // Log the field name (using BarcodeValue as a proxy) and image size in bytes.
            long imageSize = imageStream.Length;
            string logEntry = $"[{DateTime.UtcNow:O}] Field: {parameters.BarcodeValue}, ImageSize: {imageSize} bytes{Environment.NewLine}";
            File.AppendAllText("BarcodeGenerationLog.txt", logEntry);

            // Reset position for the caller.
            if (imageStream.CanSeek)
                imageStream.Position = 0;

            return imageStream;
        }

        // Generate barcode image for old-fashioned BARCODE fields.
        public Stream GetOldBarcodeImage(BarcodeParameters parameters)
        {
            return GetBarcodeImage(parameters);
        }

        // Creates a minimal 1x1 PNG image as a placeholder.
        private static Stream CreatePlaceholderImage()
        {
            // PNG data for a 1x1 transparent pixel.
            byte[] pngBytes = new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
                0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                0x54,0x78,0x9C,0x63,0x00,0x01,0x00,0x00,
                0x05,0x00,0x01,0x0D,0x0A,0x2D,0xB4,0x00,
                0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
                0x42,0x60,0x82
            };
            return new MemoryStream(pngBytes);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Assign the custom logging generator to the document.
            doc.FieldOptions.BarcodeGenerator = new LoggingBarcodeGenerator();

            // Insert a DISPLAYBARCODE field.
            FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            barcodeField.BarcodeValue = "ABC123";
            barcodeField.BarcodeType = "QR";
            barcodeField.BackgroundColor = "0xF8BD69";
            barcodeField.ForegroundColor = "0xB5413B";
            barcodeField.ErrorCorrectionLevel = "3";
            barcodeField.ScalingFactor = "250";
            barcodeField.SymbolHeight = "1000";
            barcodeField.SymbolRotation = "0";

            // Build the parameters object that the generator expects.
            BarcodeParameters parameters = new BarcodeParameters
            {
                BarcodeValue = barcodeField.BarcodeValue,
                BarcodeType = barcodeField.BarcodeType,
                BackgroundColor = barcodeField.BackgroundColor,
                ForegroundColor = barcodeField.ForegroundColor,
                ErrorCorrectionLevel = barcodeField.ErrorCorrectionLevel,
                ScalingFactor = barcodeField.ScalingFactor,
                SymbolHeight = barcodeField.SymbolHeight,
                SymbolRotation = barcodeField.SymbolRotation
            };

            // Generate the barcode image; this will also write a log entry.
            using (Stream imgStream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(parameters))
            {
                // Insert the generated image into the document.
                imgStream.Position = 0;
                builder.InsertImage(imgStream);
            }

            // Save the document.
            doc.Save("BarcodeWithLogging.docx");
        }
    }
}
