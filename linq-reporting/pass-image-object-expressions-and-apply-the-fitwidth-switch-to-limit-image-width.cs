using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code pages (required for some Aspose.Words operations).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare a simple red PNG image encoded as Base64.
        // This is a 1x1 pixel PNG with a solid red color.
        const string redPixelBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==";

        // Create the data model and assign the Base64 image string.
        var model = new ReportModel
        {
            PhotoBase64 = redPixelBase64
        };

        // Build the template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a textbox that will host the image tag.
        var textBox = builder.InsertShape(Aspose.Words.Drawing.ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);

        // LINQ Reporting tag: insert the Base64 image and limit its width with -fitWidth.
        builder.Write("<<image [model.PhotoBase64] -fitWidth>>");

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("ImageFitWidthReport.docx");
    }

    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // The Base64-encoded image that will be inserted into the document.
        public string PhotoBase64 { get; set; } = string.Empty;
    }
}
