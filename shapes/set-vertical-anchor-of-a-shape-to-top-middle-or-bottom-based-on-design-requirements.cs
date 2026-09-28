using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class SetVerticalAnchorExample
{
    // Simple enum to represent the desired vertical anchor positions.
    private enum AnchorPosition
    {
        Top,
        Center,
        Bottom
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper method to insert a floating rectangle shape with a specific vertical anchor.
        void InsertAnchoredShape(string label, AnchorPosition anchor)
        {
            // Insert a paragraph to hold the shape label.
            builder.Writeln(label);

            // Insert a rectangle shape (inline by default) and then convert it to a floating shape.
            Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
            shape.WrapType = WrapType.None;
            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

            // Determine the top position based on the requested anchor.
            // Page height is obtained from the builder's page setup (in points).
            double pageHeight = builder.PageSetup.PageHeight;
            double topPosition = anchor switch
            {
                AnchorPosition.Top => 0,
                AnchorPosition.Center => (pageHeight - shape.Height) / 2,
                AnchorPosition.Bottom => pageHeight - shape.Height,
                _ => 0
            };

            shape.Top = topPosition;
            shape.Left = 20; // A small left offset for visibility.
        }

        // Insert shapes with different vertical anchors.
        InsertAnchoredShape("Shape anchored to the top of the page:", AnchorPosition.Top);
        InsertAnchoredShape("Shape anchored to the middle of the page:", AnchorPosition.Center);
        InsertAnchoredShape("Shape anchored to the bottom of the page:", AnchorPosition.Bottom);

        // Save the document.
        string outputPath = "VerticalAnchorShapes.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
