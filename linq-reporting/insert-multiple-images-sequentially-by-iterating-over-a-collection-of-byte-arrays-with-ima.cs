using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare a simple 1x1 PNG image as a byte array
        byte[] sampleImage = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAusB9Y6vZc8AAAAASUVORK5CYII=");

        // Build the data model with several images
        ReportModel model = new ReportModel();
        for (int i = 0; i < 3; i++)
        {
            model.Images.Add(new ImageItem { Data = sampleImage });
        }

        // Create the template document programmatically
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Images Report:");
        builder.Writeln("<<foreach [img in Images]>>");

        Table table = builder.StartTable();
        builder.InsertCell();

        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 200);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Writeln("<<image [img.Data] -fitSize>>");

        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template and build the report
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(report, model, "model");

        // Save the final report
        string reportPath = "Report.docx";
        report.Save(reportPath);
    }
}

public class ReportModel
{
    public List<ImageItem> Images { get; set; } = new();
}

public class ImageItem
{
    public byte[] Data { get; set; } = Array.Empty<byte>();
}
