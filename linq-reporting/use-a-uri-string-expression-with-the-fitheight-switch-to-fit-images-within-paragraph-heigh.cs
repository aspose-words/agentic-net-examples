using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string workDir = Directory.GetCurrentDirectory();
        string outputDir = Path.Combine(workDir, "output");
        Directory.CreateDirectory(outputDir);

        // Create a minimal PNG image file
        string imagePath = Path.Combine(workDir, "sample.png");
        byte[] pngBytes = new byte[]
        {
            137,80,78,71,13,10,26,10,0,0,0,13,73,72,68,82,
            0,0,0,1,0,0,0,1,8,6,0,0,0,31,21,196,
            137,0,0,0,12,73,68,65,84,8,153,99,0,1,0,0,
            5,0,1,13,10,45,0,0,0,0,73,69,78,68,174,66,
            96,130
        };
        File.WriteAllBytes(imagePath, pngBytes);

        // Create the LINQ Reporting template
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a paragraph describing the image
        builder.Writeln("Image fitted to paragraph height using -fitHeight switch:");

        // Insert a textbox to host the image tag
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImageUri] -fitHeight>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // Prepare the data model
        ReportModel model = new ReportModel
        {
            ImageUri = Path.GetFullPath(imagePath)
        };

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document
        string outputPath = Path.Combine(outputDir, "report.docx");
        reportDoc.Save(outputPath);
    }
}

// Data model used by the LINQ Reporting engine
public class ReportModel
{
    // URI or file path to the image
    public string ImageUri { get; set; } = string.Empty;
}
