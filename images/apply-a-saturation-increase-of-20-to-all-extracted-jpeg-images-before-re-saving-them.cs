using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ImageSaturationExample
{
    public static void Main()
    {
        // Step 1: Create a sample JPEG image.
        const string inputImagePath = "sample.jpg";
        const int imgWidth = 200;
        const int imgHeight = 200;

        using (Bitmap bmp = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                using (SolidBrush brush = new SolidBrush(Color.Red))
                {
                    g.FillRectangle(brush, 20, 20, imgWidth - 40, imgHeight - 40);
                }
            }

            bmp.Save(inputImagePath, ImageFormat.Jpeg);
        }

        // Step 2: Insert the JPEG image into a new Word document.
        const string docPath = "DocumentWithImage.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // Step 3: Load the document and process each JPEG image.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int processedCount = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage) continue;

            ImageData imgData = shape.ImageData;
            string imageFormat = imgData.ImageType.ToString(); // e.g., Jpeg, Png, etc.

            if (!imageFormat.Equals("Jpeg", StringComparison.OrdinalIgnoreCase))
                continue; // Process only JPEG images.

            byte[] originalBytes = imgData.ImageBytes;
            using (MemoryStream originalStream = new MemoryStream(originalBytes))
            {
                originalStream.Position = 0;
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    using (Bitmap saturatedBitmap = new Bitmap(originalBitmap.Width, originalBitmap.Height))
                    {
                        using (Graphics graphics = Graphics.FromImage(saturatedBitmap))
                        {
                            // Build a color matrix that increases saturation by 20%.
                            float saturation = 1.2f; // 20% increase
                            float lumR = 0.3086f;
                            float lumG = 0.6094f;
                            float lumB = 0.0820f;

                            float sr = lumR * (1 - saturation) + saturation;
                            float sg = lumG * (1 - saturation) + saturation;
                            float sb = lumB * (1 - saturation) + saturation;

                            float[][] matrixElements = new float[][]
                            {
                                new float[] { sr, lumR * (1 - saturation), lumR * (1 - saturation), 0, 0 },
                                new float[] { lumG * (1 - saturation), sg, lumG * (1 - saturation), 0, 0 },
                                new float[] { lumB * (1 - saturation), lumB * (1 - saturation), sb, 0, 0 },
                                new float[] { 0, 0, 0, 1, 0 },
                                new float[] { 0, 0, 0, 0, 1 }
                            };

                            ColorMatrix colorMatrix = new ColorMatrix(matrixElements);
                            ImageAttributes imgAttributes = new ImageAttributes();
                            imgAttributes.SetColorMatrix(colorMatrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);

                            graphics.DrawImage(
                                originalBitmap,
                                new Rectangle(0, 0, originalBitmap.Width, originalBitmap.Height),
                                0,
                                0,
                                originalBitmap.Width,
                                originalBitmap.Height,
                                GraphicsUnit.Pixel,
                                imgAttributes);
                        }

                        // Save the saturated bitmap to a memory stream as JPEG.
                        using (MemoryStream saturatedStream = new MemoryStream())
                        {
                            saturatedBitmap.Save(saturatedStream, ImageFormat.Jpeg);
                            saturatedStream.Position = 0;

                            // Replace the image in the shape with the saturated version.
                            imgData.SetImage(saturatedStream);
                            processedCount++;
                        }
                    }
                }
            }
        }

        // Validate that at least one image was processed.
        if (processedCount == 0)
            throw new InvalidOperationException("No JPEG images were found to process.");

        // Step 4: Save the modified document.
        const string outputDocPath = "DocumentWithSaturatedImage.docx";
        loadedDoc.Save(outputDocPath);

        // Verify that the output file exists.
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);

        Console.WriteLine($"Processed {processedCount} JPEG image(s). Output saved to '{outputDocPath}'.");
    }
}
