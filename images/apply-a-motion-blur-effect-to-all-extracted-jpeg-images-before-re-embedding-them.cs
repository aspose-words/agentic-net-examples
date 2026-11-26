using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class MotionBlurExample
{
    public static void Main()
    {
        // Create a deterministic sample JPEG image.
        const string sampleImagePath = "sample.jpg";
        CreateSampleJpeg(sampleImagePath);

        // Insert the sample JPEG image into a new Word document.
        const string originalDocPath = "original.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(originalDocPath);

        // Load the document, apply motion blur to each JPEG image, and re‑embed.
        Document loadedDoc = new Document(originalDocPath);
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (!shape.HasImage) continue;

            // Verify the image is a JPEG by checking its header.
            byte[] imageBytes = shape.ImageData.ImageBytes;
            if (!IsJpeg(imageBytes)) continue;

            // Load the original image into a bitmap.
            using (MemoryStream originalStream = new MemoryStream(imageBytes))
            using (Bitmap originalBitmap = new Bitmap(originalStream))
            {
                // Apply a simple horizontal motion blur.
                Bitmap blurredBitmap = ApplyMotionBlur(originalBitmap, blurLength: 10, offsetX: 1);

                // Save the blurred bitmap to a memory stream as JPEG.
                using (MemoryStream blurredStream = new MemoryStream())
                {
                    blurredBitmap.Save(blurredStream, ImageFormat.Jpeg);
                    blurredStream.Position = 0;

                    // Replace the image in the shape with the blurred version.
                    shape.ImageData.SetImage(blurredStream);
                }

                blurredBitmap.Dispose();
            }
        }

        // Save the modified document.
        const string outputDocPath = "motionblur.docx";
        loadedDoc.Save(outputDocPath);

        // Validate that the output file was created.
        if (!File.Exists(outputDocPath))
            throw new InvalidOperationException("The output document was not created.");

        Console.WriteLine("Motion blur applied and document saved to: " + outputDocPath);
    }

    // Creates a deterministic sample JPEG image.
    private static void CreateSampleJpeg(string path)
    {
        const int width = 200;
        const int height = 150;
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            using (Pen pen = new Pen(Color.Blue, 5))
            {
                g.DrawEllipse(pen, 20, 20, width - 40, height - 40);
            }
            using (Brush brush = new SolidBrush(Color.Red))
            {
                g.FillRectangle(brush, 60, 60, 80, 30);
            }

            bitmap.Save(path, ImageFormat.Jpeg);
        }
    }

    // Checks JPEG signature (first two bytes 0xFF, 0xD8).
    private static bool IsJpeg(byte[] bytes)
    {
        return bytes != null && bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xD8;
    }

    // Applies a simple horizontal motion blur by drawing the image multiple times with decreasing opacity.
    private static Bitmap ApplyMotionBlur(Bitmap source, int blurLength, int offsetX)
    {
        int width = source.Width;
        int height = source.Height;
        Bitmap result = new Bitmap(width, height);
        using (Graphics g = Graphics.FromImage(result))
        {
            g.Clear(Color.Transparent);
            for (int i = 0; i < blurLength; i++)
            {
                float alpha = 1.0f - (float)i / blurLength;
                ColorMatrix cm = new ColorMatrix(new float[][]
                {
                    new float[]{1,0,0,0,0},
                    new float[]{0,1,0,0,0},
                    new float[]{0,0,1,0,0},
                    new float[]{0,0,0,alpha,0},
                    new float[]{0,0,0,0,1}
                });
                ImageAttributes ia = new ImageAttributes();
                ia.SetColorMatrix(cm);
                Rectangle destRect = new Rectangle(i * offsetX, 0, width, height);
                g.DrawImage(source, destRect, 0, 0, width, height, GraphicsUnit.Pixel, ia);
                ia.Dispose();
            }
        }
        return result;
    }
}
