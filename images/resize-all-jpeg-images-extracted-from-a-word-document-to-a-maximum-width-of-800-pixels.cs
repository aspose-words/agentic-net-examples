using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a sample JPEG image (1200x900) and save it as input.jpg
        const string sampleImagePath = "input.jpg";
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(1200, 900))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw simple text onto the image
                g.DrawString(
                    "Sample JPEG",
                    new Aspose.Drawing.Font("Arial", 48),
                    new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black),
                    new Aspose.Drawing.PointF(100, 400));
            }
            bitmap.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }

        // Create a Word document and insert the sample JPEG image twice
        const string docPath = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        builder.Writeln(); // add a line break
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Load the document for processing
        Document loadedDoc = new Document(docPath);

        // Prepare output folder
        const string outputFolder = "ResizedImages";
        Directory.CreateDirectory(outputFolder);

        // Counters for naming output files
        int imageIndex = 0;
        int resizedCount = 0;

        // Iterate over all Shape nodes in the document
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Extract the image to a memory stream
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image using Aspose.Drawing
                using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(originalStream))
                {
                    int originalWidth = originalBitmap.Width;
                    int originalHeight = originalBitmap.Height;

                    // Determine if resizing is needed
                    if (originalWidth <= 800)
                    {
                        // No resizing required; just save the original image
                        string unchangedPath = Path.Combine(outputFolder, $"image-{imageIndex}.jpg");
                        originalBitmap.Save(unchangedPath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                        imageIndex++;
                        continue;
                    }

                    // Calculate new dimensions while preserving aspect ratio
                    int newWidth = 800;
                    int newHeight = (int)(originalHeight * (800.0 / originalWidth));

                    // Create a new bitmap with the target size
                    using (Aspose.Drawing.Bitmap resizedBitmap = new Aspose.Drawing.Bitmap(newWidth, newHeight))
                    {
                        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(resizedBitmap))
                        {
                            graphics.Clear(Aspose.Drawing.Color.White);
                            graphics.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                        }

                        // Save the resized image to a file
                        string resizedPath = Path.Combine(outputFolder, $"image-{imageIndex}.jpg");
                        resizedBitmap.Save(resizedPath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                        resizedCount++;

                        // Replace the image in the document with the resized version
                        using (MemoryStream resizedStream = new MemoryStream())
                        {
                            resizedBitmap.Save(resizedStream, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                            resizedStream.Position = 0;
                            shape.ImageData.SetImage(resizedStream);
                        }
                    }

                    imageIndex++;
                }
            }
        }

        // Save the modified document
        const string outputDocPath = "output.docx";
        loadedDoc.Save(outputDocPath);

        // Validation: ensure at least one image was resized
        if (resizedCount == 0)
            throw new InvalidOperationException("No JPEG images required resizing or were found.");

        // Validation: ensure at least one resized file exists
        string[] resizedFiles = Directory.GetFiles(outputFolder, "*.jpg");
        if (resizedFiles.Length == 0)
            throw new InvalidOperationException("Resized image files were not created.");

        // Program completed successfully
    }
}
