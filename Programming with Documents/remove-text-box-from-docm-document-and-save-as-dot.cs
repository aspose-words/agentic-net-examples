using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class RemoveTextBoxAndSaveAsTemplate
{
    static void Main()
    {
        // Load the macro‑enabled DOCM document.
        // The Document(string) constructor is the approved way to create a Document from a file.
        Document doc = new Document(@"C:\Input\SourceDocument.docm");

        // Find all Shape nodes in the document (including those inside headers/footers).
        // GetChildNodes(NodeType.Shape, true) returns a live collection of all shapes.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate over the collection and remove any shape that is a text box.
        // ShapeType.TextBox identifies a text box shape.
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove(); // Remove the text box from its parent.
        }

        // Save the modified document as a DOT template.
        // Using Save(string, SaveFormat) ensures the format is explicitly set to DOT.
        doc.Save(@"C:\Output\ResultTemplate.dot", SaveFormat.Dot);
    }
}
