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

        // Define a character style with specific font attributes.
        Style charStyle = doc.Styles.Add(StyleType.Character, "MyCharStyle");
        charStyle.Font.Name = "Arial";          // Font name
        charStyle.Font.Size = 20;               // Font size in points

        // Create an Aspose.Drawing.Color and convert it to System.Drawing.Color.
        Aspose.Drawing.Color asposeColor = Aspose.Drawing.Color.Blue;
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(asposeColor.ToArgb());
        charStyle.Font.Color = sysColor;        // Font color

        // Validate that the style properties were set correctly.
        if (charStyle.Font.Name != "Arial" ||
            charStyle.Font.Size != 20 ||
            charStyle.Font.Color.ToArgb() != sysColor.ToArgb())
        {
            throw new InvalidOperationException("Style font properties were not set correctly.");
        }

        // Insert text and apply the custom character style.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Style = charStyle;
        builder.Writeln("This text uses the custom style with Arial, size 20, blue color.");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "StyledDocument.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("Failed to create the output document.", outputPath);
        }
    }
}
