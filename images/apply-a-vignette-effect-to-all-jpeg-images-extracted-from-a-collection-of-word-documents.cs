using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample JPEG image.
        // -----------------------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        Bitmap sampleBitmap = new Bitmap(imgWidth, imgHeight);
        Graphics g = Graphics.FromImage(sampleBitmap);
        g.Clear(Color.White);
        using (Pen pen = new Pen(Color.Red, 5))
        {
            g.DrawEllipse(pen, 10, 10, imgWidth - 20, imgHeight - 20);
        }
        string sampleImagePath = Path.Combine(artifactsDir, "sample.jpg");
        sampleBitmap.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        g.Dispose();
        sampleBitmap.Dispose();

        // -----------------------------------------------------------------
        // 2. Create a collection of Word documents that contain the image.
        // -----------------------------------------------------------------
        int docCount = 2;
        List<string> docPaths = new List<string>();
        for (int i = 0; i < docCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(sampleImagePath);
            string docPath = Path.Combine(artifactsDir, $"doc{i + 1}.docx");
            doc.Save(docPath);
            docPaths.Add(docPath);
        }

        // -----------------------------------------------------------------
        // 3. Extract JPEG images, apply a vignette effect and save them.
        // -----------------------------------------------------------------
        int vignetteCreated = 0;
        for (int docIndex = 0; docIndex < docPaths.Count; docIndex++)
        {
            Document doc = new Document(docPaths[docIndex]);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Save the original image.
                string ext = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string extractedPath = Path.Combine(
                    artifactsDir,
                    $"extracted_doc{docIndex + 1}_img{imageIndex}{ext}");
                shape.ImageData.Save(extractedPath);

                // Load the saved image into a bitmap.
                using (Bitmap bmp = new Bitmap(extractedPath))
                {
                    // Apply vignette using a radial gradient brush.
                    using (Graphics gfx = Graphics.FromImage(bmp))
                    {
                        Rectangle rect = new Rectangle(0, 0, bmp.Width, bmp.Height);
                        GraphicsPath path = new GraphicsPath();
                        path.AddEllipse(rect);
                        using (PathGradientBrush brush = new PathGradientBrush(path))
                        {
                            brush.CenterColor = Color.FromArgb(0, 0, 0, 0); // fully transparent center
                            brush.SurroundColors = new[] { Color.FromArgb(180, 0, 0, 0) }; // semi‑transparent black edge
                            gfx.FillRectangle(brush, rect);
                        }
                    }

                    // Save the vignetted image as JPEG.
                    string vignettePath = Path.Combine(
                        artifactsDir,
                        $"vignette_doc{docIndex + 1}_img{imageIndex}.jpg");
                    bmp.Save(vignettePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                    vignetteCreated++;
                }

                imageIndex++;
            }
        }

        // -----------------------------------------------------------------
        // 4. Validation – ensure at least one vignette image was produced.
        // -----------------------------------------------------------------
        if (vignetteCreated == 0)
            throw new InvalidOperationException("No vignette images were created.");

        // The program finishes without waiting for user input.
    }
}
