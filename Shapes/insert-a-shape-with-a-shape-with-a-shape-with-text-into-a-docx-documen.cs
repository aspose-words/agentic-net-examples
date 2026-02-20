using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create the outermost group shape that will contain all nested shapes.
        GroupShape outerGroup = new GroupShape(doc)
        {
            // Position the group shape on the page.
            Left = 100,
            Top = 100,
            Width = 400,
            Height = 300,
            // Set the coordinate system for child shapes (0,0) to (100,100).
            CoordOrigin = new System.Drawing.Point(0, 0),
            CoordSize = new System.Drawing.Size(100, 100)
        };

        // Insert the outer group shape into the document.
        builder.CurrentParagraph.AppendChild(outerGroup);

        // Create a second group shape that will be placed inside the outer group.
        GroupShape innerGroup = new GroupShape(doc)
        {
            // Position relative to the outer group's coordinate system.
            Left = 10,
            Top = 10,
            Width = 80,
            Height = 60,
            CoordOrigin = new System.Drawing.Point(0, 0),
            CoordSize = new System.Drawing.Size(100, 100)
        };

        // Add the inner group to the outer group.
        outerGroup.AppendChild(innerGroup);

        // Create a textbox shape that will hold the actual text.
        Shape textBox = new Shape(doc, ShapeType.TextBox)
        {
            // Position relative to the inner group's coordinate system.
            Left = 5,
            Top = 5,
            Width = 90,
            Height = 50,
            // Ensure the shape is displayed in front of text.
            BehindText = false,
            // Optional: give the shape a visible border.
            StrokeColor = System.Drawing.Color.Black,
            StrokeWeight = 0.5
        };

        // Add the textbox to the inner group.
        innerGroup.AppendChild(textBox);

        // Insert a paragraph into the textbox and write some text.
        Paragraph para = new Paragraph(doc);
        textBox.AppendChild(para);
        Run run = new Run(doc, "Nested shape with text");
        para.AppendChild(run);

        // Save the document to a DOCX file.
        doc.Save("NestedShapes.docx", SaveFormat.Docx);
    }
}
