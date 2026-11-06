using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare working directories.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample image file (1x1 pixel PNG).
        string imagePath = Path.Combine(outputDir, "sample.png");
        byte[] pngBytes = new byte[]
        {
            137,80,78,71,13,10,26,10,0,0,0,13,73,72,68,82,
            0,0,0,1,0,0,0,1,8,6,0,0,0,31,21,196,
            137,0,0,0,12,73,68,65,84,8,153,99,0,1,0,0,
            5,0,1,13,10,0,0,0,0,73,69,78,68,174,66,96,
            130
        };
        File.WriteAllBytes(imagePath, pngBytes);

        // Create the LINQ Reporting template.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox to host the image tag.
        var textBox = builder.InsertShape(Aspose.Words.Drawing.ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag with -fitHeight switch to scale height while preserving width proportionally.
        builder.Write("<<image [model.ImagePath] -fitHeight>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Prepare the data model.
        var model = new ReportModel
        {
            ImagePath = imagePath
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string resultPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(resultPath);

        Console.WriteLine($"Report generated: {resultPath}");
    }
}

// Data model used by the LINQ Reporting template.
public class ReportModel
{
    // Path to the image file that will be inserted into the document.
    public string ImagePath { get; set; } = string.Empty;
}
