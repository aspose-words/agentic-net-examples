using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some paragraphs.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");

        // Insert a rectangle shape to have a non‑paragraph node.
        Shape shape = new Shape(doc, ShapeType.Rectangle);
        shape.Width = 100;
        shape.Height = 50;
        builder.InsertNode(shape);

        // Iterate over all child nodes of the first paragraph using var.
        // GetChild returns a Node, but GetChildNodes is defined on CompositeNode (Paragraph inherits it).
        Paragraph firstParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        foreach (var child in firstParagraph.GetChildNodes(NodeType.Any, false))
        {
            // child is inferred as Node.
            Console.WriteLine($"Node type: {child.NodeType}");
        }

        // Iterate over the document's style collection using var.
        foreach (var style in doc.Styles)
        {
            // style is inferred as Style.
            Console.WriteLine($"Style: {style.Name}, Type: {style.Type}");
        }

        // Insert a combo box form field and iterate over its drop‑down items using var.
        builder.Writeln("Choose a value:");
        FormField combo = builder.InsertComboBox("MyCombo", new[] { "One", "Two", "Three" }, 0);
        foreach (var item in combo.DropDownItems)
        {
            // item is inferred as string.
            Console.WriteLine($"Combo box item: {item}");
        }

        // Save the document.
        doc.Save("Output.docx");
    }
}
