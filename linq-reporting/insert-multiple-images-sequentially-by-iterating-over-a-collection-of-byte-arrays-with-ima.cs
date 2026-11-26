using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

public class Program
{
    public static void Main()
    {
        // Sample images encoded as Base64 strings.
        const string base64Image1 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO6XcVYAAAAASUVORK5CYII=";
        const string base64Image2 =
            "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAEklEQVR42mP8z8Dwn4EIwDiqAAB6YQK6VgAAAABJRU5ErkJggg==";

        // Prepare the data model.
        var model = new ReportModel
        {
            Images = new List<ImageItem>
            {
                new ImageItem { Data = Convert.FromBase64String(base64Image1) },
                new ImageItem { Data = Convert.FromBase64String(base64Image2) }
            }
        };

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin the foreach block.
        builder.Writeln("<<foreach [img in Images]>>");

        // Create a table – each iteration will generate one row.
        Table table = builder.StartTable();

        // First (and only) cell of the row.
        builder.InsertCell();

        // Insert a textbox that will hold the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag that reads the byte[] from the current item.
        builder.Write("<<image [img.Data] -fitSize>>");

        // End the row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and generate the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<ImageItem> Images { get; set; } = new();
}

// Image item containing a byte array.
public class ImageItem
{
    public byte[] Data { get; set; } = Array.Empty<byte>();
}
