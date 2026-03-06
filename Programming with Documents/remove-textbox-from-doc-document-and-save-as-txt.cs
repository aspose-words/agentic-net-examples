using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class RemoveTextBoxAndSaveAsTxt
{
    static void Main()
    {
        // Paths to the source DOC/DOCX file and the destination TXT file.
        string inputPath = "input.docx";
        string outputPath = "output.txt";

        // Load the existing Word document.
        Document doc = new Document(inputPath);

        // Find all shapes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Remove every shape that is a TextBox.
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Save the modified document as plain text.
        TxtSaveOptions txtOptions = new TxtSaveOptions();
        doc.Save(outputPath, txtOptions);
    }
}
