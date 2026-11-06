using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ImageNullStreamTest
{
    // Model class with a nullable Stream property for the image.
    public class ReportModel
    {
        public Stream? ImageStream { get; set; } = null;
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a textbox shape to host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // Write the LINQ Reporting image tag that references the nullable stream.
        builder.Write("<<image [model.ImageStream]>>");

        // Prepare the model with a null image stream.
        ReportModel model = new ReportModel(); // ImageStream is already null.

        // Configure the reporting engine to use inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report. The method returns false when a graceful failure occurs.
        bool success = engine.BuildReport(doc, model, "model");

        // Verify that the engine reported a failure due to the null stream.
        if (!success)
        {
            Console.WriteLine("Test passed: Graceful failure detected for null image stream.");
        }
        else
        {
            Console.WriteLine("Test failed: Report built successfully despite null image stream.");
        }

        // Optionally save the resulting document for inspection (will contain an error placeholder).
        string outputPath = "ImageNullStreamTest_Output.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Output document saved to: {outputPath}");
    }
}
