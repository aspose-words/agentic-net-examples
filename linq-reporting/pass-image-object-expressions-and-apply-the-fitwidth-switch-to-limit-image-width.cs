using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Image data as a byte array to be used in the template.
    public byte[] Image { get; set; } = Array.Empty<byte>();
}

public class Program
{
    public static void Main()
    {
        // Working directory.
        string workDir = Directory.GetCurrentDirectory();

        // Create a simple sample image file (a 1x1 red PNG).
        string imagePath = Path.Combine(workDir, "sample.png");
        CreateSampleImage(imagePath);

        // Build the template document programmatically.
        string templatePath = Path.Combine(workDir, "template.docx");
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Prepare the data model with the image bytes.
        ReportModel model = new ReportModel
        {
            Image = File.ReadAllBytes(imagePath)
        };

        // Build the report using LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(workDir, "output.docx");
        doc.Save(outputPath);
    }

    // Writes a minimal red PNG (1x1 pixel) to the specified path.
    private static void CreateSampleImage(string path)
    {
        // Base64-encoded PNG of a single red pixel.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(path, pngBytes);
    }

    // Creates a template containing a textbox with an image tag that uses the -fitWidth switch.
    private static void CreateTemplate(string path)
    {
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a textbox that will host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Write the image tag. The expression returns a byte array from the model.
        builder.Write("<<image [model.Image] -fitWidth>>");

        // Save the template.
        template.Save(path);
    }
}
