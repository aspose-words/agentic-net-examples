using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the first paragraph that exists by default.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a run with some sample text.
        Run run = new Run(doc, "Sample text with custom font.");

        // Apply custom font formatting to the run.
        Aspose.Words.Font font = run.Font;
        font.Name = "Courier New";
        font.Size = 24;
        font.Bold = true;
        font.Color = System.Drawing.Color.Red; // Explicit System.Drawing.Color
        font.Underline = Aspose.Words.Underline.Single;

        // Add the run to the paragraph.
        paragraph.AppendChild(run);

        // Reset all font attributes of the run to their defaults.
        run.Font.ClearFormatting();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ResetFontExample.docx");
        doc.Save(outputPath);
    }
}
