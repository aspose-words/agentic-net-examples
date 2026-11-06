using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add a new paragraph style named "MyCustomStyle".
        Style style = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");

        // Set the font name for the style.
        style.Font.Name = "Arial";

        // Set the font size (points) for the style.
        style.Font.Size = 16;

        // Define a color using Aspose.Drawing.Color (red in this case).
        Aspose.Drawing.Color asposeColor = Aspose.Drawing.Color.FromArgb(255, 0, 0);

        // Convert Aspose.Drawing.Color to System.Drawing.Color.
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(asposeColor.ToArgb());

        // Apply the color to the style's font.
        style.Font.Color = sysColor;

        // Validate that the font properties were set correctly.
        if (style.Font.Name != "Arial" ||
            style.Font.Size != 16 ||
            style.Font.Color.ToArgb() != sysColor.ToArgb())
        {
            throw new InvalidOperationException("Font properties were not set correctly.");
        }

        // Create a paragraph and add a run of text.
        Paragraph paragraph = new Paragraph(doc);
        Run run = new Run(doc, "This text uses the custom style.");
        paragraph.AppendChild(run);

        // Apply the custom style to the paragraph.
        paragraph.ParagraphFormat.Style = style;

        // Add the paragraph to the document body.
        doc.FirstSection.Body.AppendChild(paragraph);

        // Save the document to a file.
        string outputPath = "StyledDocument.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved.", outputPath);
        }
    }
}
