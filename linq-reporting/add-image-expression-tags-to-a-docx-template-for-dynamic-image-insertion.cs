using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare a folder for generated files.
        string outputFolder = "output";
        Directory.CreateDirectory(outputFolder);

        // 1. Create a sample image file (a tiny red dot PNG).
        string imagePath = Path.Combine(outputFolder, "sample.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAusB9W6XK2cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // 2. Build the LINQ Reporting template programmatically.
        string templatePath = Path.Combine(outputFolder, "template.docx");
        var builder = new DocumentBuilder();

        // Insert a textbox that will hold the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // Write the image tag referencing the model property.
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template.
        builder.Document.Save(templatePath);

        // 3. Define the data model.
        var model = new ReportModel
        {
            // Use an absolute path so the engine can locate the image reliably.
            ImagePath = Path.GetFullPath(imagePath)
        };

        // 4. Load the template and build the report.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated document.
        string resultPath = Path.Combine(outputFolder, "result.docx");
        doc.Save(resultPath);
    }
}

// Public data model class required by the template.
public class ReportModel
{
    // Path to the image file that will be inserted.
    public string ImagePath { get; set; } = string.Empty;
}
