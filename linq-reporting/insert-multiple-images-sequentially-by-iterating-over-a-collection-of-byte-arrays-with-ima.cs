using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;          // Needed for Table type
using Aspose.Words.Reporting;

public class SequentialImageReport
{
    // Model class containing a collection of image data.
    public class ReportModel
    {
        public List<ImageItem> Images { get; set; } = new();
    }

    // Wrapper for a single image byte array.
    public class ImageItem
    {
        public byte[] Data { get; set; } = Array.Empty<byte>();
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin foreach over the Images collection.
        builder.Writeln("<<foreach [img in Images]>>");

        // Create a table that will hold each image in its own row.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Insert a textbox inside the cell to host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 200);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag referencing the byte[] property.
        builder.Writeln("<<image [img.Data]>>");

        // Close the row and table for this iteration.
        builder.EndRow();
        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare sample image data (two simple PNGs encoded as base64).
        // -----------------------------------------------------------------
        ReportModel model = new();

        // First sample image (a red dot).
        string redDotBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFklEQVQoU2NkYGD4z0AEYBxVSFIAAQAA//8DAF6cB9cAAAAASUVORK5CYII=";
        model.Images.Add(new ImageItem { Data = Convert.FromBase64String(redDotBase64) });

        // Second sample image (a green dot).
        string greenDotBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFklEQVQoU2NkYGD4z0AEYhRVSFIAAQAA//8DAF6cB9cAAAAASUVORK5CYII=";
        model.Images.Add(new ImageItem { Data = Convert.FromBase64String(greenDotBase64) });

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(report, model, "model");

        // Save the final report.
        const string reportPath = "Report.docx";
        report.Save(reportPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }
}
