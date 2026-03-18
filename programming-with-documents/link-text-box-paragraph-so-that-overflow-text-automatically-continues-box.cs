using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first (primary) text box.
        Shape firstBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Move the builder inside the first text box to add content.
        builder.MoveTo(firstBox.LastParagraph);
        // Write a long paragraph that will exceed the first box's capacity.
        builder.Writeln(
            "This is a long text that will overflow the first text box. " +
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
            "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi " +
            "ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit " +
            "in voluptate velit esse cillum dolore eu fugiat nulla pariatur.");

        // Move the builder back to the main document flow to insert the second box.
        builder.MoveToDocumentEnd();

        // Insert the second (linked) text box that will receive the overflow text.
        Shape secondBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Optional: adjust wrapping or positioning as needed.
        secondBox.WrapType = WrapType.None;
        secondBox.VerticalAlignment = VerticalAlignment.Top;

        // Link the first text box to the second one.
        TextBox tb1 = firstBox.TextBox;
        TextBox tb2 = secondBox.TextBox;
        if (tb1.IsValidLinkTarget(tb2))
        {
            tb1.Next = tb2; // Overflow from tb1 will continue in tb2.
        }

        // Save the resulting document.
        doc.Save("LinkedTextBoxes.docx");
    }
}
