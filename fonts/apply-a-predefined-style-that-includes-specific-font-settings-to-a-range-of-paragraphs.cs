using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a custom paragraph style with specific font settings.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 14;

        // Create a color using Aspose.Drawing.Color and convert it to System.Drawing.Color.
        Aspose.Drawing.Color asposeColor = Aspose.Drawing.Color.Blue;
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(asposeColor.ToArgb());
        customStyle.Font.Color = sysColor;

        // Apply the custom style to a range of paragraphs.
        builder.ParagraphFormat.Style = customStyle;
        builder.Writeln("First paragraph with custom style.");
        builder.Writeln("Second paragraph with custom style.");

        // Add a normal paragraph without the custom style.
        builder.ParagraphFormat.Style = doc.Styles["Normal"];
        builder.Writeln("A normal paragraph without custom style.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "StyledParagraphs.docx");
        doc.Save(outputPath);
    }
}
