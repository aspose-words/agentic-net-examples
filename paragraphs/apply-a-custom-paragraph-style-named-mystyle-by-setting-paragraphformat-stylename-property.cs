using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define a custom paragraph style named "MyStyle".
        Style myStyle = doc.Styles.Add(StyleType.Paragraph, "MyStyle");
        myStyle.Font.Name = "Arial";
        myStyle.Font.Size = 14;
        myStyle.Font.Color = Color.Blue;

        // Create a new paragraph and apply the custom style by name.
        Paragraph paragraph = new Paragraph(doc);
        paragraph.ParagraphFormat.StyleName = "MyStyle";

        // Add some text to the paragraph.
        Run run = new Run(doc, "This paragraph uses the custom style \"MyStyle\".");
        paragraph.AppendChild(run);

        // Append the paragraph to the document body.
        doc.FirstSection.Body.AppendChild(paragraph);

        // Save the document to a file.
        doc.Save("MyStyleParagraph.docx");
    }
}
