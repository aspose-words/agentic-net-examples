using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Lists;
using Aspose.Words.Themes;
using Aspose.Words.Fonts;
using Aspose.Drawing; // For Aspose.Drawing.Color

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define a custom character style.
        Style customStyle = doc.Styles.Add(StyleType.Character, "MyCustomStyle");
        // Set font name.
        customStyle.Font.Name = "Arial";
        // Set font size (points).
        customStyle.Font.Size = 24;
        // Create an Aspose.Drawing.Color and convert it to System.Drawing.Color for the Font.Color property.
        Aspose.Drawing.Color aspColor = Aspose.Drawing.Color.Blue;
        customStyle.Font.Color = System.Drawing.Color.FromArgb(aspColor.ToArgb());

        // Use DocumentBuilder to add text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is normal text.");

        // Apply the custom style to the following text.
        builder.Font.StyleName = "MyCustomStyle";
        builder.Write("This text uses the custom style.");

        // Reset formatting to default for subsequent text.
        builder.Font.ClearFormatting();
        builder.Writeln(); // Move to next line.

        // Save the document to the local file system.
        string outputPath = "StyledText.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
