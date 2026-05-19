using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Path to the image file that will be inserted into the report.
    public string ImagePath { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare a sample image file (1x1 pixel PNG) in the working folder.
        // -----------------------------------------------------------------
        const string imageFileName = "sample.png";
        byte[] imageBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK0cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imageFileName, imageBytes);

        // -----------------------------------------------------------------
        // 2. Create the data model that the LINQ Reporting engine will use.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            ImagePath = Path.GetFullPath(imageFileName)
        };

        // -----------------------------------------------------------------
        // 3. Build the template document programmatically.
        //    The image tag is placed inside a textbox so that -fitHeight can work.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a textbox with a fixed height (e.g., 150 points).
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 150);
        // Move the cursor into the textbox's first paragraph.
        builder.MoveTo(textBox.FirstParagraph);
        // Write the image tag that scales the image to fit the textbox height,
        // preserving the original width proportionally.
        builder.Write("<<image [model.ImagePath] -fitHeight>>");

        // -----------------------------------------------------------------
        // 4. Run the LINQ Reporting engine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        const string outputFileName = "output.docx";
        doc.Save(outputFileName);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputFileName)}");
    }
}
