using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Path to the image that will be placed in the footer.
    public string FooterImagePath { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare a sample image file (a tiny red dot PNG).
        // -----------------------------------------------------------------
        string imagePath = Path.Combine(Environment.CurrentDirectory, "footer.png");
        CreateSamplePng(imagePath);

        // -----------------------------------------------------------------
        // 2. Build the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        CreateTemplate(templatePath, imagePath);

        // -----------------------------------------------------------------
        // 3. Create the data model that supplies the image path.
        // -----------------------------------------------------------------
        var model = new ReportModel { FooterImagePath = imagePath };

        // -----------------------------------------------------------------
        // 4. Load the template and run the reporting engine.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
        doc.Save(outputPath);
    }

    // Creates a minimal PNG image (red 10x10 pixel) and writes it to the given path.
    private static void CreateSamplePng(string path)
    {
        // PNG binary for a 10x10 red square (base64 encoded).
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFklEQVQoU2NkYGD4z0AEYBxVSFIAAAcAAU5Z4ZcAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(path, pngBytes);
    }

    // Generates a Word document that contains a footer with an image tag.
    private static void CreateTemplate(string templatePath, string sampleImagePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Move cursor to the primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
        builder.MoveTo(textBox.FirstParagraph);

        // Write the LINQ Reporting image tag. The -fitSize switch scales the image
        // to fit the height of the container (the footer line height).
        builder.Write("<<image [model.FooterImagePath] -fitSize>>");

        // Save the template.
        doc.Save(templatePath);
    }
}
