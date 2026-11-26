using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a paragraph to hold the runs.
        Paragraph paragraph = new Paragraph(doc);

        // Create the source run with specific font formatting.
        Run sourceRun = new Run(doc, "Source text");
        sourceRun.Font.Name = "Arial";
        sourceRun.Font.Size = 16;
        sourceRun.Font.Bold = true;
        sourceRun.Font.Italic = true;
        // Use Aspose.Drawing.Color for creation, then convert to System.Drawing.Color.
        sourceRun.Font.Color = System.Drawing.Color.FromArgb(Color.Blue.ToArgb());

        // Create the target run with default formatting.
        Run targetRun = new Run(doc, "Target text");

        // Copy font formatting from sourceRun to targetRun property by property.
        targetRun.Font.Name = sourceRun.Font.Name;
        targetRun.Font.Size = sourceRun.Font.Size;
        targetRun.Font.Bold = sourceRun.Font.Bold;
        targetRun.Font.Italic = sourceRun.Font.Italic;
        targetRun.Font.Color = sourceRun.Font.Color;

        // Add both runs to the paragraph.
        paragraph.AppendChild(sourceRun);
        paragraph.AppendChild(targetRun);

        // Add the paragraph to the document body.
        doc.FirstSection.Body.AppendChild(paragraph);

        // Validate that the font properties have been copied.
        if (targetRun.Font.Name != sourceRun.Font.Name ||
            targetRun.Font.Size != sourceRun.Font.Size ||
            targetRun.Font.Bold != sourceRun.Font.Bold ||
            targetRun.Font.Italic != sourceRun.Font.Italic ||
            targetRun.Font.Color.ToArgb() != sourceRun.Font.Color.ToArgb())
        {
            throw new InvalidOperationException("Font properties were not copied correctly.");
        }

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FontCopyExample.docx");

        // Save the document.
        doc.Save(outputPath);

        // Ensure the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not created.", outputPath);
        }

        // Log success.
        Console.WriteLine("Document created successfully at: " + outputPath);
    }
}
