using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("Input.doc");

        // Collect all Shape nodes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate backwards so that removal does not affect the loop index.
        for (int i = shapes.Count - 1; i >= 0; i--)
        {
            Shape shape = (Shape)shapes[i];

            // A text box is represented by a Shape that contains a TextBox object.
            if (shape.TextBox != null)
                shape.Remove();
        }

        // Save the modified document as MHTML.
        doc.Save("Output.mhtml", SaveFormat.Mhtml);
    }
}
