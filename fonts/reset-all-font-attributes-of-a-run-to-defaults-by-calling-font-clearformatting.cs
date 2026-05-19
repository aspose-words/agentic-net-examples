using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing; // For Aspose.Drawing.Color

namespace FontResetExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a new run with some text and add it to the first paragraph.
            Run run = new Run(doc, "Hello World!");
            doc.FirstSection.Body.FirstParagraph.AppendChild(run);

            // Apply custom formatting to the run.
            Aspose.Words.Font font = run.Font;
            font.Name = "Courier New";
            font.Size = 24;
            font.Bold = true;

            // Create a color using Aspose.Drawing and convert it to System.Drawing.Color.
            Aspose.Drawing.Color asposeColor = Aspose.Drawing.Color.FromArgb(255, 0, 0); // Red
            font.Color = System.Drawing.Color.FromArgb(asposeColor.ToArgb());

            // Reset all font attributes to their defaults.
            font.ClearFormatting();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FontResetOutput.docx");
            doc.Save(outputPath);
        }
    }
}
