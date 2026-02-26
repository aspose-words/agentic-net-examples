using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class RemoveTextBoxExample
{
    static void Main()
    {
        // Path to the source DOCM file that contains a text box.
        string inputPath = @"C:\Docs\SourceDocument.docm";

        // Path where the resulting DOCX file will be saved.
        string outputPath = @"C:\Docs\ResultDocument.docx";

        // Load the existing DOCM document.
        Document doc = new Document(inputPath);

        // Retrieve all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through the collection and remove any shape that is a text box.
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
            {
                // Remove the shape (text box) from its parent node.
                shape.Remove();
            }
        }

        // Save the modified document as a macro‑free DOCX file.
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
