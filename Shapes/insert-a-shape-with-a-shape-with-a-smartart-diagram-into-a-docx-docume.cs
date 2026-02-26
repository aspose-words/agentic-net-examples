using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertSmartArt
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Load a template document that already contains a SmartArt shape.
        // The template should be prepared beforehand and placed in the same folder as the executable.
        Document smartArtTemplate = new Document("SmartArtTemplate.docx");

        // Find the first shape that has a SmartArt object.
        Shape smartArtShape = smartArtTemplate.GetChildNodes(NodeType.Shape, true)
            .Cast<Shape>()
            .FirstOrDefault(s => s.HasSmartArt);

        if (smartArtShape != null)
        {
            // Clone the SmartArt shape so it can be inserted into the new document.
            Shape clonedSmartArt = (Shape)smartArtShape.Clone(true);

            // Insert the cloned shape at the current cursor position.
            builder.InsertNode(clonedSmartArt);
        }
        else
        {
            Console.WriteLine("No SmartArt shape found in the template document.");
        }

        // Save the document with OOXML compliance that supports DML (required for SmartArt).
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save("SmartArtInserted.docx", saveOptions);
    }
}
