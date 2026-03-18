using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class TableWrapAroundImageExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating image that will be wrapped by surrounding text.
        // The image is positioned relative to the page margin and will allow text to flow around it.
        Shape image = builder.InsertImage(@"C:\Images\SampleImage.jpg");
        image.WrapType = WrapType.Square;                     // Text wraps on all sides of the image.
        image.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        image.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
        image.HorizontalAlignment = HorizontalAlignment.Left; // Align image to the left side.
        image.VerticalAlignment = VerticalAlignment.Top;
        image.BehindText = false;                            // Image appears in front of text.

        // Add a paragraph after the image to contain the table.
        builder.Writeln();

        // Start building a table that will wrap around the floating image.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Enable text wrapping for the table.
        table.TextWrapping = TextWrapping.Around;            // Wrap text around the table.
        // AllowOverlap is true by default, but we can assert it.
        if (!table.AllowOverlap)
        {
            // In practice this block will never execute because AllowOverlap defaults to true.
            throw new InvalidOperationException("Table must allow overlap for proper wrapping.");
        }

        // Adjust the distance between the table and surrounding text.
        table.AbsoluteHorizontalDistance = 10; // Points of horizontal space.
        table.AbsoluteVerticalDistance = 10;   // Points of vertical space.

        // Optionally set the anchor positions to control where the table floats.
        table.HorizontalAnchor = RelativeHorizontalPosition.Margin;
        table.VerticalAnchor = RelativeVerticalPosition.Paragraph;

        // Save the resulting document.
        doc.Save(@"C:\Output\TableWrapAroundImage.docx");
    }
}
