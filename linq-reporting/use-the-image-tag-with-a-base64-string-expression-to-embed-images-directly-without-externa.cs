using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Base64 string of a tiny PNG image (1x1 pixel, transparent).
    public string ImageBase64 { get; set; } = 
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=";
}

public class Program
{
    public static void Main()
    {
        // Prepare paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string resultPath = Path.Combine(outputDir, "Result.docx");

        // ---------- Create the LINQ Reporting template ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will hold the image.
        var textBox = builder.InsertShape(Aspose.Words.Drawing.ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Write the image tag that uses a Base64 string expression.
        builder.Write("<<image [model.ImageBase64] -fitSize>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------- Prepare the data model ----------
        var model = new ReportModel();

        // ---------- Load the template and build the report ----------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        // The root object name in the template is "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(resultPath);
    }
}
