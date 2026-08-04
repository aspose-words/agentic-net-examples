using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a custom paragraph style named "MyStyle".
        Style myStyle = doc.Styles.Add(StyleType.Paragraph, "MyStyle");
        myStyle.Font.Name = "Arial";
        myStyle.Font.Size = 14;
        myStyle.Font.Color = System.Drawing.Color.Blue;

        // Use DocumentBuilder to insert a paragraph and apply the custom style.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.StyleName = "MyStyle"; // Apply the custom style by name.
        builder.Writeln("This paragraph uses the custom style \"MyStyle\".");

        // Save the document to a file in the current directory.
        doc.Save("MyStyleParagraph.docx");
    }
}
