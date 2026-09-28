using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Path to the image file that will be inserted into the report.
    public string ImagePath { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory("output");

        // Create a simple PNG image (a red square) and save it locally.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJ" +
                                 "bWFnZVJlYWR5ccllPAAAAAlwSFlzAAAOxAAADsQBlSsOGwAAABl0RVh0Q3JlYXRpb24gVGltZQAw" +
                                 "OC8wOS8xM6V6LwAAABl0RVh0U291cmNlAEFkb2JlIEltYWdlUmVhZHlxyWUAAAAZdEVYdFNvZnR3" +
                                 "YXJlAHd3dy5pbWFnZW1hZ2UuY29t7eJ0NwAAABh0RVh0Q3JlYXRpb24gVGltZQAyMDIzLTA5LTI2" +
                                 "VDE2OjM0OjU5KzAwOjAwcZ6XWQAAABV0RVh0U291cmNlIFRpbWUAMjAyMy0wOS0yNlQxNjozNDo1" +
                                 "OSswMDowMHR6K6UAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        string imagePath = Path.GetFullPath("sample.png");
        File.WriteAllBytes(imagePath, imageBytes);

        // Create the template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a heading.
        builder.Writeln("Image Report with fitSize scaling:");

        // Insert a textbox that will contain the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);

        // Insert the LINQ Reporting image tag with -fitSize switch.
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template to disk.
        string templatePath = Path.GetFullPath("template.docx");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            ImagePath = imagePath
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine("output", "report.docx");
        reportDoc.Save(outputPath);
    }
}
