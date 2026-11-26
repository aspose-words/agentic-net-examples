using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeShapesExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare paths for the template and output documents.
            string workFolder = Path.Combine(Path.GetTempPath(), "AsposeShapesExample");
            Directory.CreateDirectory(workFolder);
            string templatePath = Path.Combine(workFolder, "Template.docx");
            string outputPath = Path.Combine(workFolder, "Output.docx");

            // Create a simple template document.
            Document templateDoc = new Document();
            DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
            templateBuilder.Writeln("This is a template document.");
            templateDoc.Save(templatePath);

            // Load the template document.
            Document doc = new Document(templatePath);
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a rectangle shape.
            builder.InsertShape(ShapeType.Rectangle, 100, 50);

            // Insert an ellipse shape.
            builder.InsertShape(ShapeType.Ellipse, 80, 80);

            // Insert a text box shape.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
            // Add text inside the text box.
            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc, "Text inside the box");
            para.AppendChild(run);
            textBox.AppendChild(para);

            // Save the modified document.
            doc.Save(outputPath);

            // Validate that the output file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not saved correctly.");

            // Optional: indicate success.
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
    }
}
