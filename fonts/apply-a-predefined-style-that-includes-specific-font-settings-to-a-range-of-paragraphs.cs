using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add five paragraphs with sample text.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Paragraph {i}");
        }

        // Define a custom paragraph style with specific font settings.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;
        // Convert Aspose.Drawing.Color to System.Drawing.Color for the Font.Color property.
        customStyle.Font.Color = System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Blue.ToArgb());

        // Apply the custom style to paragraphs 2 through 4 (zero‑based indices 1 to 3).
        for (int i = 1; i <= 3; i++)
        {
            Paragraph para = doc.FirstSection.Body.Paragraphs[i];
            para.ParagraphFormat.Style = customStyle;
        }

        // Validate that the style was applied correctly.
        bool allValid = true;
        for (int i = 1; i <= 3; i++)
        {
            Paragraph para = doc.FirstSection.Body.Paragraphs[i];
            // Each paragraph created by DocumentBuilder contains a single Run.
            Aspose.Words.Font font = para.Runs[0].Font;

            if (font.Name != "Arial" ||
                font.Size != 14 ||
                !font.Bold ||
                font.Color.ToArgb() != System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Blue.ToArgb()).ToArgb())
            {
                allValid = false;
                break;
            }
        }

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created and the style applied successfully.
        if (File.Exists(outputPath) && allValid)
        {
            Console.WriteLine("Document created and style applied successfully.");
        }
        else
        {
            Console.WriteLine("Failed to apply style or save document.");
        }
    }
}
