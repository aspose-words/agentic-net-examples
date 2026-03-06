using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // -------------------------------------------------
        // Apply Asian typography settings (East Asian layout).
        // -------------------------------------------------
        // Enable FE (Far East) layout and ensure East Asian break rules are used.
        doc.CompatibilityOptions.UseFELayout = true;
        doc.CompatibilityOptions.DoNotUseEastAsianBreakRules = false;

        // -------------------------------------------------
        // Manipulate text boxes.
        // -------------------------------------------------
        // Find the first shape that is a text box.
        Shape textBox = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (textBox != null && textBox.ShapeType == ShapeType.TextBox)
        {
            // Align the text vertically to the middle of the text box.
            textBox.TextBox.VerticalAnchor = TextBoxAnchor.Middle;

            // Replace the existing text inside the text box.
            DocumentBuilder tbBuilder = new DocumentBuilder(doc);
            tbBuilder.MoveTo(textBox.FirstParagraph);
            tbBuilder.Write("Updated text inside the text box.");
        }

        // -------------------------------------------------
        // Add a comment to the first paragraph.
        // -------------------------------------------------
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        if (firstParagraph != null)
        {
            // Create a new comment node.
            Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
            // Add comment text.
            comment.AppendChild(new Run(doc, "This is a new comment added programmatically."));
            // Attach the comment to the paragraph.
            firstParagraph.AppendChild(comment);
        }

        // -------------------------------------------------
        // Track changes for the modifications made above.
        // -------------------------------------------------
        doc.StartTrackRevisions("AutomationUser");
        // (All changes performed after StartTrackRevisions are tracked automatically.)
        doc.StopTrackRevisions();

        // -------------------------------------------------
        // Save the modified document.
        // -------------------------------------------------
        doc.Save("Output.docx");
    }
}
