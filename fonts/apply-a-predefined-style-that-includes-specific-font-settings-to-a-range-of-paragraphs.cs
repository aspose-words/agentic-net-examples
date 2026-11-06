using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing; // For Aspose.Drawing.Color

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add five paragraphs to the document.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Paragraph {i}");
        }

        // Define a new paragraph style with specific font settings.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;

        // Create a color using Aspose.Drawing.Color and convert it to System.Drawing.Color.
        Aspose.Drawing.Color asposeColor = Aspose.Drawing.Color.Blue;
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(asposeColor.ToArgb());
        customStyle.Font.Color = sysColor;

        // Apply the custom style to paragraphs 2 through 4 (zero‑based indices 1 to 3).
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;
        for (int i = 1; i <= 3; i++)
        {
            paragraphs[i].ParagraphFormat.Style = customStyle;
        }

        // Validate that the style has the expected font properties.
        if (customStyle.Font.Name != "Arial" ||
            customStyle.Font.Size != 14 ||
            !customStyle.Font.Bold ||
            customStyle.Font.Color.ToArgb() != sysColor.ToArgb())
        {
            Console.WriteLine("Font properties validation failed.");
            return;
        }

        // Save the document to disk.
        string outputPath = "StyledDocument.docx";
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
