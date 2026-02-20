using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating textbox shape (width: 300 points, height: 200 points)
        Shape shape = builder.InsertShape(ShapeType.TextBox, 300, 200);

        // Configure the shape's wrapping (floating, no text wrap)
        shape.WrapType = WrapType.None;
        shape.BehindText = false;

        // Move the builder's cursor inside the shape so that subsequent inserts go into the shape
        builder.MoveTo(shape.FirstParagraph);

        // Build a simple table inside the shape
        builder.StartTable();

        // First row (header)
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Second row (data)
        builder.InsertCell();
        builder.Writeln("Data 1");
        builder.InsertCell();
        builder.Writeln("Data 2");
        builder.EndRow();

        // Finish the table
        builder.EndTable();

        // Save the document to a DOCX file
        doc.Save("ShapeWithTable.docx");
    }
}
