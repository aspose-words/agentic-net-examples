using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a placeholder shape that will act as a SmartArt diagram.
        // We set AlternativeText to identify it as SmartArt.
        Shape placeholderSmartArt = builder.InsertShape(ShapeType.Rectangle, 300, 150);
        placeholderSmartArt.AlternativeText = "SmartArt";
        placeholderSmartArt.WrapType = WrapType.Inline;
        placeholderSmartArt.StrokeColor = Color.Black;
        placeholderSmartArt.FillColor = Color.LightBlue;

        // Save the document with the original placeholder for reference.
        string originalPath = "OriginalSmartArt.docx";
        doc.Save(originalPath);

        // Traverse all shapes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int smartArtCount = 0;

        foreach (Node node in shapeNodes)
        {
            Shape shape = (Shape)node;

            // Detect our placeholder SmartArt via AlternativeText.
            if (shape.AlternativeText == "SmartArt")
            {
                smartArtCount++;

                // Create a replacement shape (simple rectangle) with the same size.
                Shape replacement = new Shape(doc, ShapeType.Rectangle);
                replacement.Width = shape.Width;
                replacement.Height = shape.Height;
                replacement.WrapType = shape.WrapType;
                replacement.RelativeHorizontalPosition = shape.RelativeHorizontalPosition;
                replacement.RelativeVerticalPosition = shape.RelativeVerticalPosition;
                replacement.Left = shape.Left;
                replacement.Top = shape.Top;
                replacement.StrokeColor = Color.Black;
                replacement.FillColor = Color.LightGray;

                // Insert the replacement after the original shape.
                shape.ParentNode.InsertAfter(replacement, shape);
                // Remove the original placeholder SmartArt shape.
                shape.Remove();
            }
        }

        // Validate that at least one SmartArt placeholder was replaced.
        if (smartArtCount == 0)
            throw new InvalidOperationException("No SmartArt placeholders were found to replace.");

        // Save the modified document.
        string outputPath = "ReplacedSmartArt.docx";
        doc.Save(outputPath);

        // Verify that the output file exists.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
