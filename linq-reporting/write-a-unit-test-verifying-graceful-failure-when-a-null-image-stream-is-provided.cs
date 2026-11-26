using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    // Model class used as the root data source for the report.
    public class ReportModel
    {
        // ImageStream can be null to simulate missing image data.
        public Stream? ImageStream { get; set; } = null;
    }

    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document that will serve as the template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a textbox that will contain the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // LINQ Reporting image tag referencing the ImageStream property.
        builder.Write("<<image [model.ImageStream]>>");

        // Save the template to disk (required before building the report).
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // Load the template back (simulating a real-world scenario).
        Document doc = new Document(templatePath);

        // Prepare the model with a null image stream.
        ReportModel model = new ReportModel
        {
            ImageStream = null // Explicitly set to null to test graceful failure.
        };

        // Configure the reporting engine to return inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        bool success;
        try
        {
            // Build the report. The method returns false when an error occurs.
            success = engine.BuildReport(doc, model, "model");
        }
        catch (Exception ex)
        {
            // If an unexpected exception is thrown, treat it as a failure.
            Console.WriteLine($"Exception thrown: {ex.Message}");
            success = false;
        }

        // Save the resulting document (it may contain error placeholders).
        string resultPath = Path.Combine(outputDir, "Result.docx");
        doc.Save(resultPath);

        // Verify that the build failed gracefully.
        if (!success)
        {
            Console.WriteLine("Report build failed as expected due to null image stream.");
        }
        else
        {
            Console.WriteLine("Report build succeeded unexpectedly.");
        }

        // Indicate where the files are located.
        Console.WriteLine($"Template saved to: {templatePath}");
        Console.WriteLine($"Result saved to: {resultPath}");
    }
}
