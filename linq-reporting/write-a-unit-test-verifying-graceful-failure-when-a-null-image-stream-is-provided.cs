using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

#nullable enable

public class ReportModel
{
    // Image stream is intentionally left null to test graceful failure.
    public Stream? Image { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Paths for temporary files.
        const string templatePath = "template.docx";
        const string resultPath = "result.docx";

        // -------------------------------------------------
        // Create a template document with an image tag.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox to host the image tag (required by LINQ Reporting).
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.Image]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template and build the report.
        // -------------------------------------------------
        Document doc = new Document(templatePath);
        ReportModel model = new ReportModel(); // Image is null.

        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // BuildReport returns false when an error occurs and InlineErrorMessages is enabled.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save(resultPath);

        // Output the result of the operation.
        Console.WriteLine($"BuildReport succeeded: {success}");
        // Expected output: BuildReport succeeded: False
    }
}
