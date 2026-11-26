using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a sample JPEG image (1200x900) and save it locally.
        const string sampleImagePath = "sample.jpg";
        using (Bitmap bitmap = new Bitmap(1200, 900))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle for visual distinction.
                g.DrawRectangle(Pens.Black, 100, 100, 1000, 700);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Jpeg);
        }

        // Create a Word document and insert the sample JPEG image twice.
        const string inputDocPath = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        builder.Writeln(); // Add a line break.
        builder.InsertImage(sampleImagePath);
        doc.Save(inputDocPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputDocPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int resizedCount = 0;
        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Load the original image into a bitmap.
            byte[] imageBytes = shape.ImageData.ImageBytes;
            using (MemoryStream originalStream = new MemoryStream(imageBytes))
            {
                originalStream.Position = 0;
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Determine if resizing is needed.
                    if (originalBitmap.Width <= 800)
                    {
                        // No resizing required; still save a copy for validation.
                        string unchangedPath = $"resized-{imageIndex}.jpg";
                        originalBitmap.Save(unchangedPath, ImageFormat.Jpeg);
                        imageIndex++;
                        continue;
                    }

                    int newWidth = 800;
                    int newHeight = (int)((double)originalBitmap.Height * newWidth / originalBitmap.Width);

                    // Create a resized bitmap.
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics g = Graphics.FromImage(resizedBitmap))
                        {
                            g.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                        }

                        // Save resized image to a file for validation.
                        string resizedPath = $"resized-{imageIndex}.jpg";
                        resizedBitmap.Save(resizedPath, ImageFormat.Jpeg);

                        // Replace the image in the document with the resized version.
                        using (MemoryStream resizedStream = new MemoryStream())
                        {
                            resizedBitmap.Save(resizedStream, ImageFormat.Jpeg);
                            resizedStream.Position = 0;
                            shape.ImageData.SetImage(resizedStream);
                        }

                        resizedCount++;
                        imageIndex++;
                    }
                }
            }
        }

        // Save the modified document.
        const string outputDocPath = "output.docx";
        loadedDoc.Save(outputDocPath);

        // Validation: ensure at least one resized image was produced.
        if (resizedCount == 0)
        {
            throw new InvalidOperationException("No JPEG images were resized.");
        }

        // Validation: ensure the output files exist.
        if (!File.Exists(outputDocPath))
        {
            throw new FileNotFoundException("The output document was not created.", outputDocPath);
        }

        string[] resizedFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "resized-*.jpg");
        if (resizedFiles.Length == 0)
        {
            throw new InvalidOperationException("Resized image files were not created.");
        }

        // Cleanup: optional removal of temporary files can be added here.
    }
}
