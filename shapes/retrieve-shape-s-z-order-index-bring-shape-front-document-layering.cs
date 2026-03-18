using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("input.docx");

        // Retrieve all top‑level shapes in the document.
        var shapes = doc.GetChildNodes(NodeType.Shape, true)
                        .OfType<Shape>()
                        .Where(s => s.IsTopLevel)
                        .ToList();

        if (shapes.Count == 0)
        {
            Console.WriteLine("No shapes found in the document.");
            return;
        }

        // Choose the shape whose Z‑order we want to modify.
        // Here we take the first shape as an example.
        Shape targetShape = shapes[0];

        // Get the current Z‑order index of the target shape.
        int currentZOrder = targetShape.ZOrder;
        Console.WriteLine($"Current ZOrder of target shape: {currentZOrder}");

        // Find the highest Z‑order value among all shapes.
        int maxZOrder = shapes.Max(s => s.ZOrder);

        // Set the target shape's ZOrder to a value higher than any other shape.
        // This brings the shape to the front of the layering.
        targetShape.ZOrder = maxZOrder + 1;

        Console.WriteLine($"New ZOrder of target shape: {targetShape.ZOrder}");

        // Save the modified document.
        doc.Save("output.docx");
    }
}
