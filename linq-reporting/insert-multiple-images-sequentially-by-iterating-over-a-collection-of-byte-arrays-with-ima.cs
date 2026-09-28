using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Sample image data (two tiny PNG images).
        var model = new ReportModel
        {
            Images = new List<byte[]>
            {
                Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII="), // transparent 1x1
                Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==") // black 1x1
            }
        };

        // Create a template document with LINQ Reporting tags.
        const string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin foreach over the Images collection.
        builder.Writeln("<<foreach [img in Images]>>");

        // Create a table row for each image.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Insert a textbox that will hold the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [img] -fitSize>>");

        builder.EndRow();
        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "report.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Data model for the report.
public class ReportModel
{
    public List<byte[]> Images { get; set; } = new();
}
