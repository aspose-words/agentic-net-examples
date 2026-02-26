using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Attach a DocumentBuilder to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable track changes (revisions) for all subsequent edits.
        doc.StartTrackRevisions("Reviewer", DateTime.Now);

        // -----------------------------------------------------------------
        // 1. Create a list that uses an Asian font (e.g., MS Mincho for Japanese).
        // -----------------------------------------------------------------
        // Add a new list to the document (bullet list as an example).
        List list = doc.Lists.Add(ListTemplate.BulletDefault);
        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // Set Asian typography for the list items.
        builder.Font.Name = "MS Mincho"; // Japanese font; replace with appropriate Asian font if needed.
        builder.Font.Size = 12;

        // Add list items.
        builder.Writeln("項目一 – 第一項目"); // Item 1 in Japanese.
        builder.Writeln("項目二 – 第二項目"); // Item 2 in Japanese.
        builder.Writeln("項目三 – 第三項目"); // Item 3 in Japanese.

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // -----------------------------------------------------------------
        // 2. Insert a floating text box with some content.
        // -----------------------------------------------------------------
        // Insert a text box shape (floating) with specified size.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Optional: set wrap type so it behaves like a floating object.
        textBox.WrapType = WrapType.None;

        // Move the cursor inside the text box and write text.
        builder.MoveTo(textBox.FirstParagraph);
        builder.Font.Name = "SimSun"; // Chinese font as an example of Asian typography.
        builder.Font.Size = 11;
        builder.Write("这是一个文本框。"); // "This is a text box." in Chinese.

        // Return the cursor to the main story after the text box.
        builder.MoveToDocumentEnd();

        // -----------------------------------------------------------------
        // 3. Add a comment to the last paragraph of the document.
        // -----------------------------------------------------------------
        // Ensure there is a paragraph to attach the comment to.
        Paragraph targetParagraph = doc.LastSection.Body.LastParagraph;
        // Create a new comment.
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
        comment.SetText("Please review the Asian list and text box.");
        // Append the comment to the target paragraph.
        targetParagraph.AppendChild(comment);

        // -----------------------------------------------------------------
        // 4. Save the modified document.
        // -----------------------------------------------------------------
        doc.Save("Output.docx");
    }
}
