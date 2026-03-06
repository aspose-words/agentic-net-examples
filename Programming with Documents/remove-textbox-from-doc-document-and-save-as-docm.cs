using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class RemoveTextBoxAndSaveAsDocm
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = @"C:\Docs\input.doc";

        // Path for the resulting DOCM file.
        string outputPath = @"C:\Docs\output.docm";

        // Load the existing document (lifecycle: load rule).
        Document doc = new Document(inputPath);

        // Collect all shapes that are text boxes.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        for (int i = shapes.Count - 1; i >= 0; i--)
        {
            Shape shape = (Shape)shapes[i];
            // Text boxes are represented by ShapeType.TextBox.
            if (shape.ShapeType == ShapeType.TextBox)
            {
                // Remove the text box from its parent node.
                shape.Remove();
            }
        }

        // Save the modified document as a macro‑enabled DOCM file (lifecycle: save rule).
        doc.Save(outputPath, SaveFormat.Docm);
    }
}
