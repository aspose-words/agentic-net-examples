using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Folder for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create a tiny PNG image (1x1 pixel, red) from a Base64 string.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9YhZcVYAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        string imagePath = Path.Combine(workDir, "sample.png");
        File.WriteAllBytes(imagePath, imageBytes);

        // 2. Build the LINQ Reporting template programmatically.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag expects a Stream, byte[] or file path. We will supply a Stream.
        builder.Write("<<image [model.ImageStream] -fitSize>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template for reporting.
        Document loadDoc = new Document(templatePath);

        // 4. Prepare the data model. The ImageStream property returns a MemoryStream
        // that will be closed automatically by the reporting engine after insertion.
        ReportModel model = new ReportModel(imageBytes);

        // Ensure the stream is positioned at the beginning before building the report.
        model.ImageStream.Position = 0;

        // 5. Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadDoc, model, "model");

        // 6. Verify that the stream has been closed by the engine.
        bool isStreamClosed = !model.ImageStream.CanRead;
        Console.WriteLine($"Image stream closed after BuildReport: {isStreamClosed}");

        // 7. Save the final document.
        string outputPath = Path.Combine(workDir, "result.docx");
        loadDoc.Save(outputPath);
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    // The image data as a byte array (read from the sample PNG file).
    private readonly byte[] _imageData;

    // The MemoryStream that will be supplied to the image tag.
    public MemoryStream ImageStream { get; set; }

    public ReportModel(byte[] imageData)
    {
        _imageData = imageData ?? Array.Empty<byte>();
        // Initialize the stream; it will be closed automatically after report generation.
        ImageStream = new MemoryStream(_imageData);
    }
}
