using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Base working directory
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchProcessing");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputImageDir = Path.Combine(baseDir, "ExtractedImages");
        string reportPath = Path.Combine(baseDir, "ImageReport.csv");
        string sampleImagePath = Path.Combine(baseDir, "sample.png");

        // Ensure directories exist
        Directory.CreateDirectory(baseDir);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputImageDir);

        // -------------------------------------------------
        // Create a deterministic sample image (sample.png)
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // Create sample DOCX files that contain images
        // -------------------------------------------------
        int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i} - Sample content.");

            // Insert the sample image once or twice depending on the document index
            int insertCount = i; // 1,2,3 images respectively
            for (int j = 0; j < insertCount; j++)
            {
                builder.InsertImage(sampleImagePath);
                builder.Writeln(); // separate images
            }

            string docPath = Path.Combine(inputDir, $"Doc{i}.docx");
            doc.Save(docPath);
        }

        // -------------------------------------------------
        // Batch process DOCX files: extract images and collect metadata
        // -------------------------------------------------
        var records = new List<string>();
        // CSV header
        records.Add("Document,ImageFile,ImageFormat,SizeBytes,WidthPoints,HeightPoints");

        string[] docFiles = Directory.GetFiles(inputDir, "*.docx");
        foreach (string docFile in docFiles)
        {
            Document doc = new Document(docFile);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage) continue;

                // Determine image format and extension
                ImageType imgType = shape.ImageData.ImageType;
                string extension = imgType.ToString().ToLower(); // e.g., png, jpeg
                string docName = Path.GetFileNameWithoutExtension(docFile);
                string imageFileName = $"{docName}_image{imageIndex + 1}.{extension}";
                string imagePath = Path.Combine(outputImageDir, imageFileName);

                // Save the image
                shape.ImageData.Save(imagePath);

                // Gather metadata
                FileInfo fi = new FileInfo(imagePath);
                long sizeBytes = fi.Length;

                // ImageSize provides dimensions in points (double)
                var imgSize = shape.ImageData.ImageSize;
                double widthPoints = imgSize.WidthPoints;
                double heightPoints = imgSize.HeightPoints;

                // Add CSV line
                string csvLine = $"{docName},{imageFileName},{imgType},{sizeBytes},{widthPoints},{heightPoints}";
                records.Add(csvLine);

                imageIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (records.Count <= 1) // only header present
        {
            throw new InvalidOperationException("No images were extracted from the documents.");
        }

        // Write CSV report
        File.WriteAllLines(reportPath, records, Encoding.UTF8);

        // Simple validation of output files existence (optional)
        if (!File.Exists(reportPath))
        {
            throw new InvalidOperationException("CSV report was not created.");
        }

        // End of processing - program exits automatically
    }
}
