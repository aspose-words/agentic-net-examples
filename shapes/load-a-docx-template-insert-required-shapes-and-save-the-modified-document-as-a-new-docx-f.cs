using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Define temporary folder for files
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeWordsShapesExample");
        Directory.CreateDirectory(tempFolder);

        // Paths for template and output documents
        string templatePath = Path.Combine(tempFolder, "Template.docx");
        string outputPath = Path.Combine(tempFolder, "Modified.docx");

        // -------------------------------------------------
        // Step 1: Create a simple DOCX template document
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("This is the template document.");
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Step 2: Load the DOCX template
        // -------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Step 3: Insert required shapes
        // -------------------------------------------------
        // Insert a rectangle shape
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        rectangle.FillColor = Color.LightBlue;
        rectangle.StrokeColor = Color.DarkBlue;
        rectangle.StrokeWeight = 2.0; // Set line width

        // Insert a text box shape
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 60);
        textBox.FillColor = Color.LightYellow;
        textBox.StrokeColor = Color.Orange;
        textBox.StrokeWeight = 1.5; // Set line width

        // Add text inside the text box
        textBox.AppendChild(new Paragraph(doc));
        Paragraph para = (Paragraph)textBox.LastChild;
        Run run = new Run(doc, "Sample text inside a text box.");
        para.AppendChild(run);

        // -------------------------------------------------
        // Step 4: Save the modified document
        // -------------------------------------------------
        doc.Save(outputPath);

        // -------------------------------------------------
        // Step 5: Validate that the output file exists
        // -------------------------------------------------
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output document at '{outputPath}'.");
        }

        // Cleanup: (optional) delete temporary files if desired
        // File.Delete(templatePath);
        // File.Delete(outputPath);
    }
}
