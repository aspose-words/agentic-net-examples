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
        myStyle.Font.Size = 16;
        myStyle.Font.Name = "Arial";
        myStyle.Font.Color = Color.Blue;

        // Use DocumentBuilder to add a paragraph and apply the custom style.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.StyleName = "MyStyle";
        builder.Writeln("This paragraph uses the custom style MyStyle.");

        // Save the document to the local file system.
        doc.Save("MyStyleParagraph.docx");
    }
}
