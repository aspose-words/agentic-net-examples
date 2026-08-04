using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph with default formatting.
        builder.Writeln("This is normal text.");

        // ------------------------------------------------------------
        // Define a custom character style named "MyCharStyle".
        // ------------------------------------------------------------
        // Add a new character style to the document's style collection.
        Style charStyle = doc.Styles.Add(StyleType.Character, "MyCharStyle");

        // Set the desired font name.
        charStyle.Font.Name = "Arial";

        // Set the desired font size (in points).
        charStyle.Font.Size = 24;

        // Create an Aspose.Drawing.Color and convert it to System.Drawing.Color.
        Aspose.Drawing.Color aspColor = Aspose.Drawing.Color.Blue;
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());

        // Assign the color to the style's font.
        charStyle.Font.Color = sysColor;

        // ------------------------------------------------------------
        // Apply the custom style to subsequent text.
        // ------------------------------------------------------------
        // Set the builder's font style to the custom style.
        builder.Font.Style = charStyle;
        builder.Writeln("This text uses the custom style with Arial, 24pt, blue color.");

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "StyledText.docx");
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
