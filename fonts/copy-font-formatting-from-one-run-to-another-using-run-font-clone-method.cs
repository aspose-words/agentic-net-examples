using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create the source run with specific font formatting.
        Run sourceRun = new Run(doc, "Source text");
        sourceRun.Font.Name = "Courier New";
        sourceRun.Font.Size = 24;
        sourceRun.Font.Bold = true;
        sourceRun.Font.Color = System.Drawing.Color.Blue; // Fully qualified System.Drawing.Color
        paragraph.AppendChild(sourceRun);

        // Create the target run and copy the formatting from the source run.
        Run targetRun = new Run(doc, "Target text");
        targetRun.Font.Name = sourceRun.Font.Name;
        targetRun.Font.Size = sourceRun.Font.Size;
        targetRun.Font.Bold = sourceRun.Font.Bold;
        targetRun.Font.Color = sourceRun.Font.Color;
        paragraph.AppendChild(targetRun);

        // Optional validation to ensure the copy succeeded.
        if (targetRun.Font.Name != sourceRun.Font.Name ||
            targetRun.Font.Size != sourceRun.Font.Size ||
            targetRun.Font.Bold != sourceRun.Font.Bold ||
            targetRun.Font.Color.ToArgb() != sourceRun.Font.Color.ToArgb())
        {
            throw new InvalidOperationException("Font properties were not copied correctly.");
        }

        // Define output path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "CopyFontFormatting.docx");

        // Save the document.
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
