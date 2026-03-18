using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class BreakTextBoxLinkExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert four text boxes.
        Shape tbShape1 = builder.InsertShape(ShapeType.TextBox, 100, 100);
        TextBox tb1 = tbShape1.TextBox;
        builder.Writeln(); // separate boxes with a paragraph break.

        Shape tbShape2 = builder.InsertShape(ShapeType.TextBox, 100, 100);
        TextBox tb2 = tbShape2.TextBox;
        builder.Writeln();

        Shape tbShape3 = builder.InsertShape(ShapeType.TextBox, 100, 100);
        TextBox tb3 = tbShape3.TextBox;
        builder.Writeln();

        Shape tbShape4 = builder.InsertShape(ShapeType.TextBox, 100, 100);
        TextBox tb4 = tbShape4.TextBox;

        // Link the text boxes in a chain: 1 → 2 → 3 → 4.
        if (tb1.IsValidLinkTarget(tb2))
            tb1.Next = tb2;

        if (tb2.IsValidLinkTarget(tb3))
            tb2.Next = tb3;

        if (tb3.IsValidLinkTarget(tb4))
            tb3.Next = tb4;

        // At this point the flow of text goes from the first box to the fourth.
        // Break the forward link of the middle text box (tb2) so that the chain becomes:
        // 1 → 2   and   3 → 4.
        // The Previous property of tb3 points to tb2, so we call BreakForwardLink on it.
        tb3.Previous.BreakForwardLink();

        // Verify that the link has been broken.
        // tb2.Next should be null and tb3.Previous should be null.
            // (These asserts are for demonstration; they can be removed in production code.)
        System.Diagnostics.Debug.Assert(tb2.Next == null, "tb2.Next should be null after breaking the link.");
        System.Diagnostics.Debug.Assert(tb3.Previous == null, "tb3.Previous should be null after breaking the link.");

        // Save the document.
        string outputPath = "BreakForwardLinkExample.docx";
        doc.Save(outputPath);
    }
}
