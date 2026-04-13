using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the first paragraph (created by default).
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a run with custom font formatting.
        Run run = new Run(doc, "Sample text with custom formatting.");
        Aspose.Words.Font font = run.Font;
        font.Name = "Courier New";
        font.Size = 24;
        font.Color = System.Drawing.Color.Blue; // Fully qualified System.Drawing.Color
        font.Bold = true;
        font.Underline = Underline.Single;

        // Append the run to the paragraph.
        paragraph.AppendChild(run);

        // Reset all font attributes of the run to defaults.
        run.Font.ClearFormatting();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ResetFontFormatting.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + outputPath);
        }
    }
}
