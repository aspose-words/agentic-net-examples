using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph to work with.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create the source run with custom font formatting.
        Run sourceRun = new Run(doc, "Source text ");
        sourceRun.Font.Name = "Courier New";
        sourceRun.Font.Size = 24;
        sourceRun.Font.Bold = true;
        sourceRun.Font.Color = System.Drawing.Color.Blue; // Fully qualified System.Drawing.Color
        paragraph.AppendChild(sourceRun);

        // Create the target run that will receive the copied formatting.
        Run targetRun = new Run(doc, "Target text");

        // Copy font properties manually (Font.Clone does not exist).
        targetRun.Font.Name = sourceRun.Font.Name;
        targetRun.Font.Size = sourceRun.Font.Size;
        targetRun.Font.Bold = sourceRun.Font.Bold;
        targetRun.Font.Color = sourceRun.Font.Color;

        // Optional validation to ensure properties were copied correctly.
        if (targetRun.Font.Name != sourceRun.Font.Name ||
            targetRun.Font.Size != sourceRun.Font.Size ||
            targetRun.Font.Bold != sourceRun.Font.Bold ||
            targetRun.Font.Color.ToArgb() != sourceRun.Font.Color.ToArgb())
        {
            Console.WriteLine("Font properties were not copied correctly.");
        }

        paragraph.AppendChild(targetRun);

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CopyFontFormatting.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
