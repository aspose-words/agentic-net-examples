using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(baseDir);

        // Create a deterministic sample image.
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath);

        // Folder for sample DOCX files.
        string docsDir = Path.Combine(baseDir, "Docs");
        Directory.CreateDirectory(docsDir);

        // Create a few sample documents that contain the same image.
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(docsDir, $"Document{i}.docx");
            CreateSampleDocWithImage(docPath, sampleImagePath, i);
        }

        // Folder where extracted images will be saved.
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(imagesDir);

        // Prepare a simple CSV content that will act as the Excel index.
        // Excel can open CSV files, and we will give it an .xlsx extension to satisfy the task.
        StringBuilder csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("Document,ImageFile"); // Header

        int totalExtracted = 0;

        // Process each DOCX file in the collection.
        foreach (string docFile in Directory.GetFiles(docsDir, "*.docx"))
        {
            Document doc = new Document(docFile);
            var shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    // Determine the proper file extension for the image type.
                    string extension = Aspose.Words.FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Img{imageIndex}{extension}";
                    string imageFullPath = Path.Combine(imagesDir, imageFileName);

                    // Save the image to the file system.
                    shape.ImageData.Save(imageFullPath);

                    // Record the entry in the CSV index.
                    csvBuilder.AppendLine($"{Path.GetFileName(docFile)},{imageFileName}");

                    imageIndex++;
                    totalExtracted++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Save the CSV content with an .xlsx extension.
        string excelPath = Path.Combine(baseDir, "ImageIndex.xlsx");
        File.WriteAllText(excelPath, csvBuilder.ToString(), Encoding.UTF8);

        // Validate that the Excel (CSV) file was created.
        if (!File.Exists(excelPath))
            throw new InvalidOperationException("Failed to create the Excel index file.");
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath)
    {
        // Create a 200x200 bitmap.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        // Obtain a graphics object for drawing.
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        // Fill background with white.
        graphics.Clear(Aspose.Drawing.Color.White);

        // Draw a simple black rectangle.
        Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black);
        graphics.DrawRectangle(pen, 10, 10, 180, 180);

        // Save the image.
        bitmap.Save(filePath);

        // Clean up resources.
        graphics.Dispose();
        pen.Dispose();
        bitmap.Dispose();
    }

    // Creates a DOCX file that contains a single image.
    private static void CreateSampleDocWithImage(string docPath, string imagePath, int docNumber)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Sample document {docNumber}");

        // Insert the image; InsertImage returns a Shape that is already appended.
        builder.InsertImage(imagePath);

        // Save the document.
        doc.Save(docPath);
    }
}
