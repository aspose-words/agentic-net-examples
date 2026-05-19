using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 1. Insert a floating rectangle shape.
        // -----------------------------------------------------------------
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,                     // shape type
            RelativeHorizontalPosition.Page, 100,    // left position (points) relative to page
            RelativeVerticalPosition.Page, 100,      // top position (points) relative to page
            100,                                      // width (points)
            50,                                       // height (points)
            WrapType.None);                           // floating (no text wrap)

        // Verify that the shape is floating.
        if (shape.IsInline)
            throw new Exception("The shape should be floating after insertion.");

        // -----------------------------------------------------------------
        // 2. Convert the floating shape to an inline shape.
        // -----------------------------------------------------------------
        shape.WrapType = WrapType.Inline;               // make it inline
        // Inline shapes ignore positioning, so reset related properties.
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Margin;
        shape.Left = 0;
        shape.Top = 0;

        // Verify conversion.
        if (!shape.IsInline)
            throw new Exception("The shape should be inline after conversion.");

        // -----------------------------------------------------------------
        // 3. Revert the inline shape back to a floating shape.
        // -----------------------------------------------------------------
        shape.WrapType = WrapType.None;                 // back to floating
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.Left = 150;                               // new left position
        shape.Top = 150;                                // new top position

        // Verify final state.
        if (shape.IsInline)
            throw new Exception("The shape should be floating after reverting.");

        // Save the resulting document.
        doc.Save("FloatingInlineShape.docx");
    }
}
