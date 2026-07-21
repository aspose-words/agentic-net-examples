using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Drawing;               // Aspose.Drawing for bitmap creation
using Aspose.Drawing.Imaging;      // For ImageFormat
using Newtonsoft.Json;             // Required package as per specification

public class ExtractVideoFrameImages
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample image that will act as a video frame thumbnail
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "video_frame.png");
        const int imgWidth = 200;
        const int imgHeight = 200;

        // Use fully‑qualified Aspose.Drawing types to avoid ambiguity
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.LightBlue);
            using (var pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.DarkBlue, 5))
            {
                graphics.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
            }
            using (var font = new Aspose.Drawing.Font("Arial", 20))
            using (var brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
            {
                graphics.DrawString("Video Frame", font, brush, new Aspose.Drawing.PointF(20, imgHeight / 2 - 15));
            }
            bitmap.Save(sampleImagePath);
        }

        // ---------------------------------------------------------------
        // 2. Create a DOCX document and insert the sample image (as a placeholder for a video frame)
        // ---------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        doc.Save(docPath);

        // ---------------------------------------------------------------
        // 3. Load the document (simulating a real DOCX with embedded video frames)
        // ---------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions();
        Document loadedDoc = new Document(docPath, loadOptions);

        // ---------------------------------------------------------------
        // 4. Extract all images from shapes (including video frame thumbnails) and save as high‑resolution PNG
        // ---------------------------------------------------------------
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Determine a suitable file name with the correct extension
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            // Force PNG output for high‑resolution requirement
            if (!extension.Equals(".png", StringComparison.OrdinalIgnoreCase))
                extension = ".png";

            string outFile = Path.Combine(artifactsDir, $"extracted_{extractedCount}{extension}");

            // Save the image data directly; this preserves the original image bytes.
            shape.ImageData.Save(outFile);
            extractedCount++;
        }

        // ---------------------------------------------------------------
        // 5. Validation – ensure at least one image was extracted
        // ---------------------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }
}
