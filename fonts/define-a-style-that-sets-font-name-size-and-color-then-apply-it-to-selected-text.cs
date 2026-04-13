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

        // Define a custom character style.
        Style charStyle = doc.Styles.Add(StyleType.Character, "MyCharStyle");

        // Set font name and size.
        charStyle.Font.Name = "Arial";
        charStyle.Font.Size = 24;

        // Create an Aspose.Drawing.Color and convert it to System.Drawing.Color.
        Aspose.Drawing.Color asposeColor = Aspose.Drawing.Color.FromArgb(255, 0, 0); // Red.
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(asposeColor.ToArgb());

        // Assign the color to the style's font.
        charStyle.Font.Color = sysColor;

        // Validate that the style's font properties are set correctly.
        if (charStyle.Font.Name != "Arial" ||
            charStyle.Font.Size != 24 ||
            charStyle.Font.Color.ToArgb() != sysColor.ToArgb())
        {
            throw new InvalidOperationException("Style font properties validation failed.");
        }

        // Write normal text without the custom style.
        builder.Writeln("This is normal text.");

        // Apply the custom style to the following text.
        builder.Font.StyleName = "MyCharStyle";
        builder.Writeln("This text uses the custom style.");

        // Reset to the default style for any further text.
        builder.Font.StyleName = "Default Paragraph Font";

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "StyledDocument.docx");
        doc.Save(outputPath);

        // Verify that the file was created successfully.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not created.", outputPath);
        }

        // Indicate success.
        Console.WriteLine("Document created successfully at: " + outputPath);
    }
}
