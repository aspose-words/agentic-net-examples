using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Drawing;

namespace FontClearFormattingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Get the first paragraph of the document (created by default).
            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

            // Create a run with some text.
            Run run = new Run(doc, "Formatted text");

            // Apply custom font formatting.
            Aspose.Words.Font font = run.Font;
            font.Name = "Courier New";
            font.Size = 24;
            // Use Aspose.Drawing.Color to define the color, then convert to System.Drawing.Color.
            font.Color = System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Blue.ToArgb());

            // Add the run to the paragraph.
            paragraph.AppendChild(run);

            // Reset all font attributes of the run to their defaults.
            run.Font.ClearFormatting();

            // Save the document to the local file system.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
            doc.Save(outputPath);
        }
    }
}
