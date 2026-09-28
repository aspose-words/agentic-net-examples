using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create output directory.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create the template document with an image tag inside a textbox.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox to host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImageStream]>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Prepare the model with a null image stream.
        ReportModel model = new ReportModel();

        // Configure the reporting engine to use inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated document.
        string resultPath = Path.Combine(outputDir, "result.docx");
        reportDoc.Save(resultPath);

        // Verify that the engine reported failure due to the null image stream.
        if (!success)
        {
            Console.WriteLine("Test Passed: BuildReport returned false for null image stream.");
        }
        else
        {
            Console.WriteLine("Test Failed: BuildReport succeeded unexpectedly.");
        }
    }

    // Model class used by the LINQ Reporting template.
    public class ReportModel
    {
        // Intentionally null to simulate a missing image.
        public Stream? ImageStream { get; set; } = null;
    }
}
