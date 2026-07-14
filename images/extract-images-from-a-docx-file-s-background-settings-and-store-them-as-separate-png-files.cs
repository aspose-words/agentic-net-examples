using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image that will be used as the document background.
        const string backgroundImagePath = "bg.png";
        CreateSampleImage(backgroundImagePath);

        // Build a DOCX document and assign the created image as its background shape.
        const string docPath = "sample.docx";
        CreateDocumentWithBackground(docPath, backgroundImagePath);

        // Load the document and extract the background image, saving it as a separate PNG file.
        const string extractedImagePath = "extracted_background.png";
        ExtractBackgroundImage(docPath, extractedImagePath);

        // Validate that the extraction succeeded.
        if (!File.Exists(extractedImagePath))
            throw new InvalidOperationException($"Failed to create '{extractedImagePath}'.");

        // Indicate successful completion.
        Console.WriteLine("Background image extracted to: " + extractedImagePath);
    }

    private static void CreateSampleImage(string filePath)
    {
        // 200x200 white canvas with a red ellipse.
        using (var bitmap = new Bitmap(200, 200))
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            using (var brush = new SolidBrush(Color.Red))
            {
                graphics.FillEllipse(brush, 20, 20, 160, 160);
            }

            // Save the image as PNG.
            bitmap.Save(filePath);
        }
    }

    private static void CreateDocumentWithBackground(string docFilePath, string imageFilePath)
    {
        // Create a new empty document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Create a rectangle shape that will serve as the background.
        var backgroundShape = new Shape(doc, ShapeType.Rectangle);

        // Assign the previously created image to the shape.
        backgroundShape.ImageData.SetImage(imageFilePath);

        // Optionally adjust shape size to match the page.
        backgroundShape.Width = doc.FirstSection.PageSetup.PageWidth;
        backgroundShape.Height = doc.FirstSection.PageSetup.PageHeight;

        // Set the shape as the document's background.
        doc.BackgroundShape = backgroundShape;

        // Save the document.
        doc.Save(docFilePath);
    }

    private static void ExtractBackgroundImage(string docFilePath, string outputImagePath)
    {
        // Load the document that contains a background shape.
        var doc = new Document(docFilePath);

        // Retrieve the background shape.
        var backgroundShape = doc.BackgroundShape;

        // Ensure the shape exists and actually contains an image.
        if (backgroundShape == null || !backgroundShape.HasImage)
            throw new InvalidOperationException("The document does not contain a background image.");

        // Determine the appropriate file extension based on the image type.
        string extension = FileFormatUtil.ImageTypeToExtension(backgroundShape.ImageData.ImageType);

        // Ensure we save as PNG regardless of original format.
        // If the original is not PNG, we still save using the original extension.
        // For this task we expect PNG, but we handle any format safely.
        string finalPath = Path.ChangeExtension(outputImagePath, extension);

        // Save the image data to the file system.
        backgroundShape.ImageData.Save(finalPath);
    }
}
