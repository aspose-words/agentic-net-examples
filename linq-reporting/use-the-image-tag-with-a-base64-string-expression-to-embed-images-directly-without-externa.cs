using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Sample data model with a Base64 image string (without the data URI prefix).
        ReportModel model = new()
        {
            // 1x1 red PNG encoded as Base64.
            ImageBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg=="
        };

        // Create a blank template document.
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Insert a textbox that will host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // LINQ Reporting image tag – the expression returns a Base64 string.
        builder.Write("<<image [model.ImageBase64] -fitSize>>");

        // Populate the template with the data model.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("ReportWithBase64Image.docx");
    }
}

// Public data model required by the template.
public class ReportModel
{
    // Base64‑encoded image data (no data URI prefix).
    public string ImageBase64 { get; set; } = string.Empty;
}
