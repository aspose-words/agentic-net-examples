using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;

namespace FontStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Define a custom paragraph style with specific font settings.
            Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
            customStyle.Font.Name = "Arial";
            customStyle.Font.Size = 14;
            customStyle.Font.Bold = true;
            // Convert Aspose.Drawing.Color to System.Drawing.Color.
            customStyle.Font.Color = System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Blue.ToArgb());

            // Add three paragraphs to the document.
            Paragraph para1 = new Paragraph(doc);
            para1.AppendChild(new Run(doc, "First paragraph"));
            doc.FirstSection.Body.AppendChild(para1);

            Paragraph para2 = new Paragraph(doc);
            para2.AppendChild(new Run(doc, "Second paragraph"));
            doc.FirstSection.Body.AppendChild(para2);

            Paragraph para3 = new Paragraph(doc);
            para3.AppendChild(new Run(doc, "Third paragraph"));
            doc.FirstSection.Body.AppendChild(para3);

            // Apply the custom style to the first two paragraphs.
            para1.ParagraphFormat.Style = customStyle;
            para2.ParagraphFormat.Style = customStyle;

            // Validate that the style was applied correctly.
            ValidateFontSettings(para1, "Arial", 14, true, System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Blue.ToArgb()));
            ValidateFontSettings(para2, "Arial", 14, true, System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Blue.ToArgb()));
            // The third paragraph retains default formatting (no validation needed).

            // Save the document.
            string outputPath = "Output.docx";
            doc.Save(outputPath);

            // Ensure the output file exists.
            if (File.Exists(outputPath))
            {
                Console.WriteLine("Document saved successfully to " + Path.GetFullPath(outputPath));
            }
            else
            {
                Console.WriteLine("Failed to save the document.");
            }
        }

        private static void ValidateFontSettings(Paragraph paragraph, string expectedName, double expectedSize, bool expectedBold, System.Drawing.Color expectedColor)
        {
            // Use fully qualified Aspose.Words.Font type.
            Aspose.Words.Font font = paragraph.Runs[0].Font;
            bool nameMatches = string.Equals(font.Name, expectedName, StringComparison.OrdinalIgnoreCase);
            bool sizeMatches = Math.Abs(font.Size - expectedSize) < 0.01;
            bool boldMatches = font.Bold == expectedBold;
            bool colorMatches = font.Color.ToArgb() == expectedColor.ToArgb();

            if (!nameMatches || !sizeMatches || !boldMatches || !colorMatches)
            {
                Console.WriteLine("Font validation failed for paragraph: \"" + paragraph.GetText().Trim() + "\"");
            }
        }
    }
}
