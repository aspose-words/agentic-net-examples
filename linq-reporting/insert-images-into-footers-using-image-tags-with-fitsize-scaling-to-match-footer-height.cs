using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string ImagePath { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Prepare sample image.
        const string imageFileName = "sample.png";
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9YVhZV8AAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imageFileName, imageBytes);

        // Create template document with footer containing an image tag.
        const string templateFileName = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Move to primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert a textbox to host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 50);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template.
        templateDoc.Save(templateFileName);

        // Load the template for reporting.
        Document doc = new Document(templateFileName);

        // Prepare model with image path.
        ReportModel model = new ReportModel
        {
            ImagePath = Path.GetFullPath(imageFileName)
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        const string outputFileName = "output.docx";
        doc.Save(outputFileName);
    }
}
