using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputFile = "input.doc";

        // Path where the resulting PostScript file will be saved.
        string outputFile = "output.ps";

        // Load the existing document.
        Document doc = new Document(inputFile);

        // Find all Shape nodes (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate backwards so that removal does not affect the loop index.
        for (int i = shapes.Count - 1; i >= 0; i--)
        {
            Shape shape = (Shape)shapes[i];

            // Remove the shape if it is a TextBox.
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Configure PostScript save options.
        PsSaveOptions psOptions = new PsSaveOptions
        {
            // Explicitly set the format; not strictly required but clarifies intent.
            SaveFormat = SaveFormat.Ps
        };

        // Save the modified document as a PostScript file.
        doc.Save(outputFile, psOptions);
    }
}
