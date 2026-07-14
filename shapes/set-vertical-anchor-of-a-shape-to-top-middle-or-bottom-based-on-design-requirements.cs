using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class SetShapeVerticalAnchorExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "VerticalAnchorDemo.docx");

        // Helper to insert a textbox with a specific vertical anchor.
        void InsertTextBox(double left, double top, TextBoxAnchor anchor, string label)
        {
            // Insert a floating textbox.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
            textBox.WrapType = WrapType.None;
            textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBox.Left = left;
            textBox.Top = top;

            // Set the vertical anchor of the text inside the textbox.
            textBox.TextBox.VerticalAnchor = anchor;

            // Add a paragraph with label text inside the textbox.
            builder.MoveTo(textBox.FirstParagraph);
            builder.Font.Size = 12;
            builder.Font.Bold = true;
            builder.Write(label);
        }

        // Insert three textboxes with Top, Middle, and Bottom vertical anchors.
        InsertTextBox(100, 100, TextBoxAnchor.Top, "Top Anchor");
        InsertTextBox(350, 100, TextBoxAnchor.Middle, "Middle Anchor");
        InsertTextBox(600, 100, TextBoxAnchor.Bottom, "Bottom Anchor");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
