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

        // Create a run with some text.
        Run run = new Run(doc, "Sample text with custom formatting.");

        // Apply custom font formatting.
        Aspose.Words.Font font = run.Font;
        font.Name = "Courier New";
        font.Size = 24;
        font.Bold = true;

        // Set the font color using Aspose.Drawing.Color and convert to System.Drawing.Color.
        Aspose.Drawing.Color aspColor = Aspose.Drawing.Color.Red;
        font.Color = System.Drawing.Color.FromArgb(aspColor.ToArgb());

        // Append the run to the first paragraph of the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Save the document with custom formatting.
        string beforePath = Path.Combine(Directory.GetCurrentDirectory(), "BeforeClear.docx");
        doc.Save(beforePath);

        // Reset all font attributes of the run to defaults.
        run.Font.ClearFormatting();

        // Save the document after clearing formatting.
        string afterPath = Path.Combine(Directory.GetCurrentDirectory(), "AfterClear.docx");
        doc.Save(afterPath);
    }
}
