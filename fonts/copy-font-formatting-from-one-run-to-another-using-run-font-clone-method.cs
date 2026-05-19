using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing; // For Aspose.Drawing.Color

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a paragraph to hold the runs.
        Paragraph para = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(para);

        // Create the source run with specific font formatting.
        Run sourceRun = new Run(doc, "Source text");
        sourceRun.Font.Name = "Arial";
        sourceRun.Font.Size = 24;
        sourceRun.Font.Bold = true;
        // Font.Color expects System.Drawing.Color; convert from Aspose.Drawing.Color.
        sourceRun.Font.Color = System.Drawing.Color.FromArgb(Color.Blue.ToArgb());
        para.AppendChild(sourceRun);

        // Create the target run whose formatting will be copied.
        Run targetRun = new Run(doc, "Target text");
        para.AppendChild(targetRun);

        // Copy font properties from the source run to the target run.
        targetRun.Font.Name = sourceRun.Font.Name;
        targetRun.Font.Size = sourceRun.Font.Size;
        targetRun.Font.Bold = sourceRun.Font.Bold;
        targetRun.Font.Color = System.Drawing.Color.FromArgb(sourceRun.Font.Color.ToArgb());

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CopyFontFormatting.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
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
