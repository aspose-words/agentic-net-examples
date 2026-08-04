using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Newtonsoft.Json;

namespace FontResetExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Add a paragraph to the first section.
            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

            // Create a run with some sample text.
            Run run = new Run(doc, "Sample text with custom formatting.");
            paragraph.AppendChild(run);

            // Apply various font attributes to the run.
            Aspose.Words.Font font = run.Font;
            font.Name = "Courier New";
            font.Size = 24;
            font.Bold = true;
            font.Italic = true;
            font.Underline = Underline.Single;
            font.Color = System.Drawing.Color.Blue; // Explicit System.Drawing.Color

            // Reset all font attributes to their defaults.
            font.ClearFormatting();

            // Define output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FontResetOutput.docx");

            // Save the document.
            doc.Save(outputPath);
        }
    }
}
