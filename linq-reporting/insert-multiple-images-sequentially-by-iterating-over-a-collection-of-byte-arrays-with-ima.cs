using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Needed for the Table class

public class Program
{
    // Model class that will be passed to the LINQ Reporting engine.
    public class ReportModel
    {
        // Collection of image data as byte arrays.
        public List<byte[]> Images { get; set; } = new();
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare sample image data (three tiny PNG images encoded in Base64).
        // -----------------------------------------------------------------
        var model = new ReportModel();

        // 1x1 pixel transparent PNG.
        const string base64Png1 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9YV4cZcAAAAASUVORK5CYII=";
        // 1x1 pixel red PNG.
        const string base64Png2 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==";
        // 1x1 pixel green PNG.
        const string base64Png3 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==";

        model.Images.Add(Convert.FromBase64String(base64Png1));
        model.Images.Add(Convert.FromBase64String(base64Png2));
        model.Images.Add(Convert.FromBase64String(base64Png3));

        // -----------------------------------------------------------------
        // 2. Build the template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Images inserted sequentially:");

        // Start the foreach block that iterates over the Images collection.
        builder.Writeln("<<foreach [img in Images]>>");

        // Create a table where each row will contain one image.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Insert a textbox that will host each image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag that consumes the current byte[] (img) and fits it to the textbox size.
        builder.Write("<<image [img] -fitSize>>");

        // Close the cell/row/table – the whole block will be repeated for each image.
        builder.EndRow();
        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before building the report).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and generate the report using LINQ Reporting.
        // -----------------------------------------------------------------
        var loadedTemplate = new Document(templatePath);
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // Build the report; the root object name must match the tag reference ("model").
        engine.BuildReport(loadedTemplate, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the final document.
        // -----------------------------------------------------------------
        const string outputPath = "Report.docx";
        loadedTemplate.Save(outputPath);
    }
}
