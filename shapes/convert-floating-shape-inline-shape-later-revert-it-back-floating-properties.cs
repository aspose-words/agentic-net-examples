using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class ConvertShapeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating image (initially floating).
        Shape floatingShape = builder.InsertImage(@"Images\Sample.png");
        floatingShape.WrapType = WrapType.None;               // Floating shape – no text wrap.
        floatingShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        floatingShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        floatingShape.Left = 100;   // Position 100 points from the left edge of the page.
        floatingShape.Top = 100;    // Position 100 points from the top edge of the page.

        // Save the document with the floating shape.
        doc.Save("FloatingShape.docx");

        // ------------------------------------------------------------
        // Convert the floating shape to an inline shape.
        // Inline shapes are treated as characters inside a paragraph,
        // so we set the WrapType to Inline.
        floatingShape.WrapType = WrapType.Inline;

        // When a shape becomes inline, its position properties (Left, Top, etc.)
        // are ignored. The shape will now flow with the surrounding text.
        // Save the document to see the inline version.
        doc.Save("InlineShape.docx");

        // ------------------------------------------------------------
        // Revert the shape back to floating.
        // To make it floating again, set WrapType to a non‑inline value,
        // e.g., WrapType.None (no text wrap) or WrapType.Square, etc.
        floatingShape.WrapType = WrapType.None;

        // Restore the position properties if needed.
        floatingShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        floatingShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        floatingShape.Left = 150;   // New horizontal position.
        floatingShape.Top = 150;    // New vertical position.

        // Save the document with the shape back to floating.
        doc.Save("FloatingShapeAgain.docx");
    }
}
