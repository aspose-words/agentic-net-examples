using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace AsposeWordsExample
{
    public class Program
    {
        // Macro‑like routine that inserts a table, a paragraph and a linked text box.
        private static void InsertElements(DocumentBuilder builder)
        {
            // ----- Insert a simple 2x2 table -----
            builder.StartTable();

            // First row
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            // Second row
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();

            builder.EndTable();

            // ----- Insert a paragraph after the table -----
            builder.Writeln("This paragraph follows the table.");

            // ----- Insert two text boxes and link them -----
            // First text box (source)
            Shape shape1 = builder.InsertShape(ShapeType.TextBox, 200, 100);
            shape1.WrapType = WrapType.None;
            shape1.HorizontalAlignment = HorizontalAlignment.Center;
            shape1.VerticalAlignment = VerticalAlignment.Top;

            // Add text to the first text box
            builder.MoveTo(shape1.FirstParagraph);
            builder.Write("First linked text box");

            // Second text box (target)
            Shape shape2 = builder.InsertShape(ShapeType.TextBox, 200, 100);
            shape2.WrapType = WrapType.None;
            shape2.HorizontalAlignment = HorizontalAlignment.Center;
            shape2.VerticalAlignment = VerticalAlignment.Top;

            // Add text to the second text box
            builder.MoveTo(shape2.FirstParagraph);
            builder.Write("Second linked text box");

            // Link the first text box to the second one if the link is valid
            TextBox tb1 = shape1.TextBox;
            TextBox tb2 = shape2.TextBox;
            if (tb1.IsValidLinkTarget(tb2))
                tb1.Next = tb2;
        }

        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a DocumentBuilder attached to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Perform the combined insertion.
            InsertElements(builder);

            // Save the resulting document.
            doc.Save("LinkedElements.docx");
        }
    }
}
