using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class ReportModel
{
    // Base64 string of a 1x1 pixel PNG image.
    public string ImageBase64 { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Paths for template and output documents.
        const string templatePath = "template.docx";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create the template document with an image tag inside a textbox.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox shape that will hold the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // LINQ Reporting image tag referencing the model's Base64 string.
        builder.Write("<<image [model.ImageBase64] -fitSize>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template for report generation.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Sample data model with a Base64-encoded PNG image.
        ReportModel model = new ReportModel
        {
            ImageBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BFwAE/wJ/lK5XAAAAAElFTkSuQmCC"
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save(outputPath);
    }
}
