using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Paths to the source and destination documents.
        string inputPath = "input.docx";
        string outputPath = "output.docx";

        // Load the existing DOCX document.
        Document doc = new Document(inputPath);

        // Enable tracking of revisions (track changes).
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Create a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 1. Insert a list using an Asian font (e.g., MS Mincho).
        // -------------------------------------------------
        builder.Font.Name = "MS Mincho";               // Asian typography.
        builder.ListFormat.ApplyNumberDefault();       // Start a numbered list.
        builder.Writeln("項目一");                     // List item 1.
        builder.Writeln("項目二");                     // List item 2.
        builder.ListFormat.RemoveNumbers();            // End the list.

        // -------------------------------------------------
        // 2. Insert a text box shape and add Asian text inside it.
        // -------------------------------------------------
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Font.Name = "MS Mincho";
        builder.Write("テキストボックスの内容");

        // -------------------------------------------------
        // 3. Add a comment to the current paragraph.
        // -------------------------------------------------
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
        comment.SetText("これはコメントです。");
        builder.CurrentParagraph.AppendChild(comment);

        // -------------------------------------------------
        // 4. Stop tracking revisions.
        // -------------------------------------------------
        doc.StopTrackRevisions();

        // Save the modified document.
        doc.Save(outputPath);
    }
}
