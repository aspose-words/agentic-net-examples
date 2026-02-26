using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace GroupShapeExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ---------------------------------------------------------------------
            // NOTE: The DocumentBuilder.StartGroupShape / EndGroupShape methods are
            // available only in recent versions of Aspose.Words (v23.5+). If you are
            // using an older version, create the GroupShape manually as shown
            // below.
            // ---------------------------------------------------------------------

            // Create a GroupShape and insert it at the builder's current position.
            GroupShape group = new GroupShape(doc);
            builder.InsertNode(group);

            // Move the builder's cursor inside the newly created group so that any
            // subsequent nodes are added as children of the group.
            builder.MoveTo(group);

            // -----------------------------------------------------------------
            // Add a rectangle shape to the group.
            // -----------------------------------------------------------------
            Shape rect = new Shape(doc, ShapeType.Rectangle)
            {
                Width = 150,
                Height = 100,
                FillColor = Color.LightBlue,
                Stroke = { Color = Color.DarkBlue }
            };
            group.AppendChild(rect);

            // -----------------------------------------------------------------
            // Add a star shape to the group.
            // -----------------------------------------------------------------
            Shape star = new Shape(doc, ShapeType.Star)
            {
                Width = 80,
                Height = 80,
                FillColor = Color.Yellow,
                Stroke = { Color = Color.Orange }
            };
            group.AppendChild(star);

            // (Optional) Move the builder back to the document body if you need to
            // continue adding content outside the group.
            builder.MoveToDocumentEnd();

            // Save the document as DOCX.
            doc.Save("GroupShape_StartGroupShape.docx");
        }
    }
}
