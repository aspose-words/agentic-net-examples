using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with sample text.
        builder.Writeln("This paragraph will have a custom style applied.");

        // Define a custom paragraph style named "MyStyle".
        Style myStyle = doc.Styles.Add(StyleType.Paragraph, "MyStyle");
        myStyle.Font.Color = Color.Blue;
        myStyle.Font.Size = 16;
        myStyle.Font.Bold = true;

        // Apply the custom style to the paragraph by setting StyleName.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        paragraph.ParagraphFormat.StyleName = "MyStyle";

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
