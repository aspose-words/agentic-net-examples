using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class ReportModel
{
    // URL of the image to be inserted into the report.
    public string ImageUrl { get; set; } = "https://via.placeholder.com/150";
}

public class Program
{
    public static void Main()
    {
        // Step 1: Create the LINQ Reporting template.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Insert a textbox that will hold the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // LINQ Reporting image tag referencing the model's ImageUrl property.
        builder.Write("<<image [model.ImageUrl] -fitSize>>");

        // Save the template to disk.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Step 2: Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var model = new ReportModel();

        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Step 3: Save the generated report.
        const string outputPath = "output.docx";
        reportDoc.Save(outputPath);
    }
}
