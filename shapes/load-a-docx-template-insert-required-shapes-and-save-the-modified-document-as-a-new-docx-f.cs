using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ShapeInsertionExample
{
    public static void Main()
    {
        // Define directories and file paths.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        string templatePath = Path.Combine(artifactsDir, "Template.docx");
        string resultPath = Path.Combine(artifactsDir, "Result.docx");
        string imagePath = Path.Combine(artifactsDir, "SampleImage.png");

        // -------------------------------------------------
        // 1. Create a simple placeholder PNG image.
        // -------------------------------------------------
        // This is a 1x1 pixel transparent PNG encoded in Base64.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK9cAAAAASUVORK5CYII=";
        File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

        // -------------------------------------------------
        // 2. Create a DOCX template and save it to disk.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("This is a template document.");
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 3. Load the template document.
        // -------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 4. Insert a rectangle shape.
        // -------------------------------------------------
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        rectangle.FillColor = System.Drawing.Color.Yellow;
        rectangle.Stroke.Color = System.Drawing.Color.Black;

        // -------------------------------------------------
        // 5. Insert an ellipse shape.
        // -------------------------------------------------
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 100, 100);
        ellipse.FillColor = System.Drawing.Color.LightGreen;
        ellipse.Stroke.Color = System.Drawing.Color.Blue;

        // -------------------------------------------------
        // 6. Insert a text box shape with some text.
        // -------------------------------------------------
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 80);
        textBox.FillColor = System.Drawing.Color.LightGray;

        // Ensure the text box contains a paragraph, then add a run.
        Paragraph tbParagraph = textBox.FirstParagraph ?? new Paragraph(doc);
        if (textBox.FirstParagraph == null)
            textBox.AppendChild(tbParagraph);

        Run tbRun = new Run(doc, "Hello from TextBox!");
        tbParagraph.AppendChild(tbRun);

        // -------------------------------------------------
        // 7. Insert an image shape using the generated image.
        // -------------------------------------------------
        Shape imageShape = builder.InsertImage(imagePath);
        imageShape.Width = 100;
        imageShape.Height = 100;

        // -------------------------------------------------
        // 8. Save the modified document.
        // -------------------------------------------------
        doc.Save(resultPath);

        // -------------------------------------------------
        // 9. Validate that the output file exists.
        // -------------------------------------------------
        if (!File.Exists(resultPath))
            throw new Exception("The result document was not created.");
    }
}
