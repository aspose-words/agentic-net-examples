using System;
using System.IO;
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

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the custom style to the next paragraph by setting StyleName.
        builder.ParagraphFormat.StyleName = "MyStyle";

        // Write a paragraph that will use the custom style.
        builder.Writeln("This paragraph uses the custom style \"MyStyle\".");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MyStyleParagraph.docx");
        doc.Save(outputPath);
    }
}
