using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    private const string SampleImagePath = "sample.png";
    private const string DownloadedImagePath = "downloaded.png";
    private const string TemplatePath = "template.docx";
    private const string OutputPdfPath = "output.pdf";

    public static void Main()
    {
        // Step 1: Create a deterministic sample image.
        CreateSampleImage();

        // Step 2: Simulate a REST API call and retrieve the image.
        byte[] imageBytes = GetImageFromApi();

        // Step 3: Save the retrieved image locally.
        File.WriteAllBytes(DownloadedImagePath, imageBytes);

        // Step 4: Create a DOCX template with a bookmark where the image will be inserted.
        CreateTemplateDocument();

        // Step 5: Load the template and insert the image.
        InsertImageIntoDocument();

        // Step 6: Validate that the PDF was created.
        ValidateOutput();

        Console.WriteLine("PDF created successfully at: " + Path.GetFullPath(OutputPdfPath));
    }

    private static void CreateSampleImage()
    {
        // Create a 200x100 PNG with a red ellipse on a white background.
        Bitmap bitmap = new Bitmap(200, 100);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        using (SolidBrush brush = new SolidBrush(Color.Red))
        {
            graphics.FillEllipse(brush, 10, 10, 180, 80);
        }

        bitmap.Save(SampleImagePath, ImageFormat.Png);
        graphics.Dispose();
        bitmap.Dispose();

        if (!File.Exists(SampleImagePath) || new FileInfo(SampleImagePath).Length == 0)
            throw new Exception("Failed to create sample image.");
    }

    // Simulates a REST API that would return the image as a base‑64 JSON payload.
    // For the purpose of this example we simply read the locally created image file.
    private static byte[] GetImageFromApi()
    {
        if (!File.Exists(SampleImagePath))
            throw new FileNotFoundException("Sample image not found.", SampleImagePath);

        return File.ReadAllBytes(SampleImagePath);
    }

    private static void CreateTemplateDocument()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a DOCX template.");
        builder.StartBookmark("ImageHere");
        builder.EndBookmark("ImageHere");
        builder.Writeln("Text after the image placeholder.");
        doc.Save(TemplatePath);

        if (!File.Exists(TemplatePath) || new FileInfo(TemplatePath).Length == 0)
            throw new Exception("Failed to create template document.");
    }

    private static void InsertImageIntoDocument()
    {
        Document doc = new Document(TemplatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToBookmark("ImageHere");
        builder.InsertImage(DownloadedImagePath);
        doc.Save(OutputPdfPath, SaveFormat.Pdf);
    }

    private static void ValidateOutput()
    {
        if (!File.Exists(OutputPdfPath) || new FileInfo(OutputPdfPath).Length == 0)
            throw new Exception("PDF output was not created successfully.");
    }
}
