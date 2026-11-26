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

        // Define a new character style named "MyCustomStyle".
        Style charStyle = doc.Styles.Add(StyleType.Character, "MyCustomStyle");

        // Set font name and size.
        charStyle.Font.Name = "Arial";
        charStyle.Font.Size = 16;

        // Create an Aspose.Drawing.Color and convert it to System.Drawing.Color.
        Aspose.Drawing.Color aspColor = Aspose.Drawing.Color.Blue;
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());

        // Assign the color to the style's font.
        charStyle.Font.Color = sysColor;

        // Add normal text.
        builder.Writeln("This text is normal.");

        // Add text that will use the custom style.
        builder.Font.StyleName = "MyCustomStyle";
        builder.Writeln("This text uses MyCustomStyle.");

        // Validate that the style's font properties were set correctly.
        Aspose.Words.Font styleFont = charStyle.Font;
        if (styleFont.Name != "Arial" ||
            styleFont.Size != 16 ||
            styleFont.Color.ToArgb() != sysColor.ToArgb())
        {
            throw new Exception("Font properties were not set correctly on the style.");
        }

        // Save the document.
        string outputPath = "StyledDocument.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output file was not created.");
        }
    }
}
