using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing; // Only for Aspose.Drawing.Color

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define a custom character style.
        Style charStyle = doc.Styles.Add(StyleType.Character, "MyCharStyle");

        // Set font name and size.
        charStyle.Font.Name = "Arial";
        charStyle.Font.Size = 24;

        // Create an Aspose.Drawing.Color and convert it to System.Drawing.Color.
        Aspose.Drawing.Color asposeColor = Aspose.Drawing.Color.Blue;
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(asposeColor.ToArgb());

        // Apply the color to the style.
        charStyle.Font.Color = sysColor;

        // Use DocumentBuilder to add text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph uses the default style.");

        // Apply the custom character style to the next paragraph.
        builder.Font.Style = charStyle;
        builder.Writeln("This paragraph uses the custom character style.");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "StyledDocument.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }

        // Validate that the style properties are set correctly.
        bool styleValid = charStyle.Font.Name == "Arial"
                          && charStyle.Font.Size == 24
                          && charStyle.Font.Color.ToArgb() == sysColor.ToArgb();

        Console.WriteLine(styleValid ? "Style properties validated." : "Style validation failed.");
    }
}
