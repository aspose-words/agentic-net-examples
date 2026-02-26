using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("Input.docx");

        // Start tracking all subsequent changes as revisions.
        doc.StartTrackRevisions("Automated");

        // ---------- Asian typography settings ----------
        // Ensure East Asian line‑break rules are applied.
        doc.CompatibilityOptions.DoNotUseEastAsianBreakRules = false;
        // Allow vertical alignment inside text boxes (relevant for Asian layout).
        doc.CompatibilityOptions.DoNotVertAlignInTxbx = false;

        // ---------- Text box manipulation ----------
        // Find the first text box in the document; if none exists, create one.
        Shape textBox = FindFirstTextBox(doc);
        if (textBox == null)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);
            textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("Initial textbox content");
        }

        // Replace the contents of the text box.
        textBox.FirstParagraph.Runs.Clear();
        textBox.FirstParagraph.AppendChild(new Run(doc, "Updated textbox content"));

        // ---------- Comment manipulation ----------
        // Add a new comment to the first paragraph.
        Paragraph firstPara = doc.FirstSection.Body.FirstParagraph;
        Comment newComment = new Comment(doc, "John Doe", "JD", DateTime.Now);
        newComment.SetText("This is a new comment.");
        firstPara.AppendChild(newComment);

        // If a comment already exists, edit its text.
        Comment existingComment = (Comment)doc.GetChild(NodeType.Comment, 0, true);
        if (existingComment != null)
        {
            existingComment.SetText("Edited comment text.");
        }

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Helper: returns the first textbox shape found in the document, or null.
    private static Shape FindFirstTextBox(Document doc)
    {
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                return shape;
        }
        return null;
    }
}
