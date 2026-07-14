using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph and a placeholder shape that will act as a SmartArt diagram.
        builder.Writeln("Below is a SmartArt diagram:");
        Shape smartArtPlaceholder = builder.InsertShape(ShapeType.Rectangle, 300, 200);
        smartArtPlaceholder.AlternativeText = "SmartArt"; // Mark this shape as SmartArt.
        smartArtPlaceholder.FillColor = Color.LightBlue;
        smartArtPlaceholder.StrokeColor = Color.DarkBlue;
        smartArtPlaceholder.TextPath.Text = "SmartArt Diagram";
        smartArtPlaceholder.TextPath.FontFamily = "Arial";
        smartArtPlaceholder.TextPath.Bold = true;
        smartArtPlaceholder.TextPath.FitPath = true;

        // Save the original document.
        string originalPath = "Original.docx";
        doc.Save(originalPath);

        // Traverse all shapes and replace the ones marked as SmartArt.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            // Detect SmartArt using the AlternativeText marker.
            if (shape.AlternativeText == "SmartArt")
            {
                // Create a replacement rectangle shape.
                Shape replacement = new Shape(doc, ShapeType.Rectangle)
                {
                    Width = shape.Width,
                    Height = shape.Height,
                    WrapType = WrapType.Inline,
                    FillColor = Color.LightGray,
                    StrokeColor = Color.Black
                };

                // Add descriptive text.
                replacement.TextPath.Text = "Replaced Diagram";
                replacement.TextPath.FontFamily = "Arial";
                replacement.TextPath.Bold = true;
                replacement.TextPath.FitPath = true;

                // Insert the replacement after the original shape and remove the original.
                shape.ParentNode.InsertAfter(replacement, shape);
                shape.Remove();
            }
        }

        // Save the modified document.
        string modifiedPath = "Modified.docx";
        doc.Save(modifiedPath);

        // Verify that both files were created.
        if (!File.Exists(originalPath) || !File.Exists(modifiedPath))
        {
            throw new Exception("Failed to save the output documents.");
        }

        Console.WriteLine("Documents created successfully:");
        Console.WriteLine($"- Original: {Path.GetFullPath(originalPath)}");
        Console.WriteLine($"- Modified: {Path.GetFullPath(modifiedPath)}");
    }
}
