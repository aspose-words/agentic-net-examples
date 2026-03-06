using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class RemoveTextBoxAndSaveAsXps
{
    static void Main()
    {
        // Path to the folder that contains the input document and where the output will be saved.
        string docsPath = @"C:\Docs\";

        // Load the existing DOC document.
        Document doc = new Document(docsPath + "Input.docx");

        // Find all shape nodes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Remove every shape that is a TextBox.
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Prepare XPS save options (default options are sufficient).
        XpsSaveOptions saveOptions = new XpsSaveOptions();

        // Save the modified document as XPS.
        doc.Save(docsPath + "Output.xps", saveOptions);
    }
}
