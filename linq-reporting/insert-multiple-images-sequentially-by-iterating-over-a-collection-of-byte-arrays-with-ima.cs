using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class ImageItem
{
    public byte[] Data { get; set; } = Array.Empty<byte>();
}

public class ReportModel
{
    public List<ImageItem> Images { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample images as byte arrays (1x1 PNG pixels).
        var redPixelBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        var bluePixelBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/lcKcAAAAAElFTkSuQmCC";

        var model = new ReportModel();
        model.Images.Add(new ImageItem { Data = Convert.FromBase64String(redPixelBase64) });
        model.Images.Add(new ImageItem { Data = Convert.FromBase64String(bluePixelBase64) });

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template.
        // -----------------------------------------------------------------
        var templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin foreach over the Images collection.
        builder.Writeln("<<foreach [img in Images]>>");

        // Table that will hold each image in its own row.
        var table = builder.StartTable();
        builder.InsertCell();

        // TextBox acts as a container for the image tag (required by the rules).
        var textBox = builder.InsertShape(ShapeType.TextBox, 200, 200);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Writeln("<<image [img.Data]>>");

        // Close the row and table for this iteration.
        builder.EndRow();
        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        reportDoc.Save("Report.docx");
    }
}
