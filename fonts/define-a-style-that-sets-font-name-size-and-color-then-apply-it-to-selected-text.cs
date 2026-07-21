using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Lists;
using Aspose.Words.Themes;
using Aspose.Drawing; // For Aspose.Drawing.Color

public class Program
{
    public static void Main()
    {
        // Output file path
        string outputPath = "StyledDocument.docx";

        // Create a new blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a custom character style with specific font settings
        Style customStyle = doc.Styles.Add(StyleType.Character, "MyCharStyle");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 24;

        // Set font color using Aspose.Drawing.Color and convert to System.Drawing.Color
        Aspose.Drawing.Color aspColor = Aspose.Drawing.Color.Blue;
        customStyle.Font.Color = System.Drawing.Color.FromArgb(aspColor.ToArgb());

        // Validate that the style properties were set correctly
        if (customStyle.Font.Name != "Arial" ||
            customStyle.Font.Size != 24 ||
            customStyle.Font.Color.ToArgb() != System.Drawing.Color.FromArgb(aspColor.ToArgb()).ToArgb())
        {
            throw new InvalidOperationException("Failed to set style font properties.");
        }

        // Write some normal text
        builder.Writeln("This is normal text.");

        // Apply the custom style to the following text
        builder.Font.StyleName = "MyCharStyle";
        builder.Writeln("This text uses the custom style with Arial, 24pt, blue color.");

        // Save the document
        doc.Save(outputPath);

        // Ensure the file was created
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");
    }
}
