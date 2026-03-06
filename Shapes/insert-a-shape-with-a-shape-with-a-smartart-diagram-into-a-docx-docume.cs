using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertSmartArtExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Load a template document that already contains a SmartArt shape.
        // The template should be placed in the same folder as the executable or provide a full path.
        Document template = new Document("SmartArtTemplate.docx");

        // Find the first shape that has a SmartArt object.
        Shape smartArtShape = template.GetChildNodes(NodeType.Shape, true)
                                      .Cast<Shape>()
                                      .FirstOrDefault(s => s.HasSmartArt);

        if (smartArtShape != null)
        {
            // Clone the SmartArt shape so it can be inserted into the new document.
            Shape clonedSmartArt = (Shape)smartArtShape.Clone(true);

            // Insert the cloned shape at the current cursor position.
            builder.InsertNode(clonedSmartArt);

            // Ensure the SmartArt drawing is rendered correctly.
            clonedSmartArt.UpdateSmartArtDrawing();
        }
        else
        {
            Console.WriteLine("No SmartArt shape found in the template document.");
        }

        // Save the resulting document.
        doc.Save("SmartArtInserted.docx", SaveFormat.Docx);
    }
}
