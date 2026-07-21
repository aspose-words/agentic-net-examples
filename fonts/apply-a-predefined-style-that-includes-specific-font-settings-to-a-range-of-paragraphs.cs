using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;
using Aspose.Drawing;

public class ApplyStyleExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a custom paragraph style with specific font settings.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        // Set font name, size and bold attribute.
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;
        // Create an Aspose.Drawing.Color and convert it to System.Drawing.Color.
        Aspose.Drawing.Color asposeColor = Aspose.Drawing.Color.Blue;
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(asposeColor.ToArgb());
        // Apply the color to the style's font.
        customStyle.Font.Color = sysColor;

        // Apply the custom style to a range of paragraphs.
        builder.ParagraphFormat.Style = customStyle;
        builder.Writeln("First paragraph using the custom style.");
        builder.Writeln("Second paragraph also using the custom style.");
        builder.Writeln("Third paragraph still using the custom style.");

        // Add a normal paragraph without the custom style for comparison.
        builder.ParagraphFormat.Style = doc.Styles["Normal"];
        builder.Writeln("A normal paragraph without the custom style.");

        // Determine an output path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "StyledDocument.docx");

        // Save the document to the file system.
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
