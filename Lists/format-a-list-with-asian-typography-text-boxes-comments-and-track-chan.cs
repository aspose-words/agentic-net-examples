using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("input.docx");

        // Enable tracking of revisions (track changes).
        doc.StartTrackRevisions("Reviewer");

        // Create a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Format a list using Asian typography (e.g., Japanese font).
        // -------------------------------------------------
        builder.Font.Name = "MS Mincho";   // Asian font.
        builder.Font.Size = 12;

        // Apply a default numbered list style.
        builder.ListFormat.ApplyNumberDefault();

        // Add list items.
        builder.Writeln("項目一");
        builder.Writeln("項目二");
        builder.Writeln("項目三");

        // Remove list formatting after the list is done.
        builder.ListFormat.RemoveNumbers();

        // -------------------------------------------------
        // Insert a text box and write Asian text inside it.
        // -------------------------------------------------
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Move the cursor into the text box's first paragraph.
        builder.MoveTo(textBox.FirstParagraph);
        builder.Font.Name = "MS Mincho";
        builder.Write("テキストボックスの内容");

        // -------------------------------------------------
        // Add a comment to the paragraph following the text box.
        // -------------------------------------------------
        Paragraph para = builder.CurrentParagraph;
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("これはコメントです。");
        para.AppendChild(comment);

        // -------------------------------------------------
        // Stop tracking revisions.
        // -------------------------------------------------
        doc.StopTrackRevisions();

        // Save the modified document.
        doc.Save("output.docx");
    }
}
