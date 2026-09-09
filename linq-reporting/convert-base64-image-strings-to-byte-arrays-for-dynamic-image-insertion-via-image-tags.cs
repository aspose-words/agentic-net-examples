using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class ReportModel
{
    // Title to display in the report.
    public string Title { get; set; } = "Sample Image Report";

    // Base64-encoded PNG image (1x1 transparent pixel).
    public string ImageBase64 { get; set; } = 
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK9cAAAAASUVORK5CYII=";

    // Converts the Base64 string to a byte array for the image tag.
    public byte[] ImageBytes => Convert.FromBase64String(ImageBase64);
}

public class Program
{
    public static void Main()
    {
        // Prepare the data source.
        var model = new ReportModel();

        // Create a blank document that will serve as the template.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a textbox to host the image tag (required by LINQ Reporting).
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag expects a byte[]; the expression returns model.ImageBytes.
        builder.Write("<<image [model.ImageBytes] -fitSize>>");

        // Add a title below the image.
        builder.Writeln();
        builder.Writeln("<<[model.Title]>>");

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}
