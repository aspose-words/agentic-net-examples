using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class LinqReportingImageFitHeightExample
{
    public static void Main()
    {
        // Create a simple PNG image file that will be used in the report.
        const string imageFileName = "sample.png";
        CreateSamplePng(imageFileName);

        // Prepare the data model.
        ReportModel model = new()
        {
            ImageUri = Path.GetFullPath(imageFileName)
        };

        // Create a new blank document that will serve as the template.
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Add a paragraph that will contain a textbox for the image.
        builder.Writeln("Below is an image that fits the paragraph height:");
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);

        // Insert the LINQ Reporting image tag with the -fitHeight switch.
        // The expression returns a URI string (the full path to the PNG file).
        builder.Write("<<image [model.ImageUri] -fitHeight>>");

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        const string outputFileName = "ReportWithFitHeightImage.docx";
        doc.Save(outputFileName);
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputFileName)}");
    }

    // Helper method to create a tiny PNG file (1x1 pixel, red) if it does not exist.
    private static void CreateSamplePng(string filePath)
    {
        if (File.Exists(filePath))
            return;

        // PNG data for a 1x1 red pixel.
        byte[] pngData = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");
        File.WriteAllBytes(filePath, pngData);
    }

    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // URI (file path) of the image to be displayed.
        public string ImageUri { get; set; } = string.Empty;
    }
}
