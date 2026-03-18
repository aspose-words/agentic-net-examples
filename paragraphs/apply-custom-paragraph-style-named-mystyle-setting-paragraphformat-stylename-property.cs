using System;
using System.Drawing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define a custom paragraph style called "MyStyle".
        Style myStyle = doc.Styles.Add(StyleType.Paragraph, "MyStyle");
        myStyle.Font.Name = "Times New Roman";
        myStyle.Font.Size = 14;
        myStyle.Font.Color = Color.Blue;

        // Create a paragraph and apply the custom style via StyleName.
        Paragraph para = new Paragraph(doc);
        para.ParagraphFormat.StyleName = "MyStyle";

        // Add some text to the paragraph.
        Run run = new Run(doc, "This paragraph uses MyStyle.");
        para.AppendChild(run);

        // Insert the paragraph into the document body.
        doc.FirstSection.Body.AppendChild(para);

        // Save the document to a file.
        doc.Save("MyStyleParagraph.docx");
    }
}
