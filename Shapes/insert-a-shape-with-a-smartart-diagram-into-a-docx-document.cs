using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Load a template document that already contains a SmartArt shape
        string templatePath = @"C:\Templates\SmartArtTemplate.docx";
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Template not found: {templatePath}");
            return;
        }
        Document smartArtTemplate = new Document(templatePath);

        // Find the first shape that has a SmartArt object in the template
        Shape smartArtShape = smartArtTemplate.GetChildNodes(NodeType.Shape, true)
            .Cast<Shape>()
            .FirstOrDefault(s => s.HasSmartArt);

        // If a SmartArt shape was found, clone it and insert it into the new document
        if (smartArtShape != null)
        {
            // Clone the shape (deep clone to copy all child nodes)
            Shape clonedShape = (Shape)smartArtShape.Clone(true);

            // Ensure the SmartArt drawing is up‑to‑date (necessary when the pre‑rendered drawing is missing)
            clonedShape.UpdateSmartArtDrawing();

            // Insert the cloned SmartArt shape at the current cursor position
            builder.InsertNode(clonedShape);
        }
        else
        {
            Console.WriteLine("No SmartArt shape found in the template.");
        }

        // Save the resulting document
        string outputPath = @"C:\Output\DocumentWithSmartArt.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
