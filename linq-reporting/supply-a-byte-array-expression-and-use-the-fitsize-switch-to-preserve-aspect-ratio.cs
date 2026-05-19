using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a textbox that will hold the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // LINQ Reporting tag that uses a byte[] expression and the -fitSize switch.
        builder.Write("<<image [model.ImageBytes] -fitSize>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document report = new Document(templatePath);

        // Prepare the data model with a sample PNG image as a byte array.
        ReportModel model = new ReportModel();

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // Save the generated report.
        report.Save("Report.docx");
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    // A 1x1 pixel transparent PNG image encoded in Base64.
    public byte[] ImageBytes { get; set; } = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/6VbZAAAAAElFTkSuQmCC");
}
