using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document (lifecycle: load)
        Document doc = new Document("Input.docx");

        // Enable tracking of revisions (track changes)
        doc.StartTrackRevisions("AutomatedUser", DateTime.Now);

        // Create a DocumentBuilder attached to the loaded document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 1. Insert a bulleted list with Asian typography
        // -------------------------------------------------
        // Use a built‑in bullet list style
        builder.ListFormat.ApplyBulletDefault();
        // Example Asian characters (Japanese, Chinese, Korean)
        builder.Writeln("日本語の項目 1"); // Japanese
        builder.Writeln("中文項目 2");   // Chinese
        builder.Writeln("한국어 항목 3"); // Korean
        // End the list
        builder.ListFormat.RemoveNumbers();

        // -------------------------------------------------
        // 2. Insert a floating text box
        // -------------------------------------------------
        // Insert a text box shape (floating)
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Optional: set wrap type so it floats over text
        textBox.WrapType = WrapType.None;
        // Move the cursor inside the text box and add some text
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("これはテキストボックスです。"); // Sample Japanese text

        // -------------------------------------------------
        // 3. Add a comment to the last paragraph of the document
        // -------------------------------------------------
        // Retrieve the last paragraph (where the cursor currently is)
        Paragraph targetParagraph = builder.CurrentParagraph;
        // Create a comment authored by "Reviewer"
        Comment comment = new Comment(doc, "Reviewer", "R", DateTime.Now);
        comment.SetText("Please review this paragraph.");
        // Append the comment to the paragraph
        targetParagraph.AppendChild(comment);

        // -------------------------------------------------
        // Save the modified document (lifecycle: save)
        // -------------------------------------------------
        doc.Save("Output.docx");
    }
}
