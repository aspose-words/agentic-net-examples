using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeFontsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a custom paragraph style named "MyCustomStyle".
            Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");

            // Define the font settings for the style.
            Aspose.Words.Font styleFont = customStyle.Font;
            styleFont.Name = "Arial";
            styleFont.Size = 14.0;
            styleFont.Bold = true;
            styleFont.Color = System.Drawing.Color.DarkGreen; // Fully qualified System.Drawing.Color

            // Apply the custom style to a range of paragraphs.
            builder.ParagraphFormat.Style = customStyle;
            builder.Writeln("This paragraph uses the custom style.");
            builder.Writeln("Another paragraph with the same style.");
            builder.Writeln("Yet another styled paragraph.");

            // Revert to the default style for subsequent paragraphs.
            builder.ParagraphFormat.Style = doc.Styles["Normal"];
            builder.Writeln("This paragraph uses the normal style.");
            builder.Writeln("Back to the default formatting.");

            // Define the output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "StyledParagraphs.docx");

            // Save the document.
            doc.Save(outputPath);

            // Validate that the file was created.
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
}
