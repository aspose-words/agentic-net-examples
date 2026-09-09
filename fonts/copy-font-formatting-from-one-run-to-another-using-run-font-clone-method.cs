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

        // Ensure the document has at least one paragraph.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;

        // Create the source run with custom font formatting.
        Run sourceRun = new Run(doc, "Source text");
        // Set font name, size, bold and color.
        sourceRun.Font.Name = "Courier New";
        sourceRun.Font.Size = 24;
        sourceRun.Font.Bold = true;
        // Convert Aspose.Drawing.Color to System.Drawing.Color as required by the API.
        sourceRun.Font.Color = System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Red.ToArgb());

        // Append the source run to the paragraph.
        para.AppendChild(sourceRun);

        // Create the destination run with default formatting.
        Run destRun = new Run(doc, "Destination text");
        para.AppendChild(destRun);

        // Copy the font properties from the source run to the destination run.
        destRun.Font.Name = sourceRun.Font.Name;
        destRun.Font.Size = sourceRun.Font.Size;
        destRun.Font.Bold = sourceRun.Font.Bold;
        destRun.Font.Color = sourceRun.Font.Color;

        // Validate that the font properties were copied correctly.
        if (destRun.Font.Name != sourceRun.Font.Name ||
            destRun.Font.Size != sourceRun.Font.Size ||
            destRun.Font.Bold != sourceRun.Font.Bold ||
            destRun.Font.Color.ToArgb() != sourceRun.Font.Color.ToArgb())
        {
            throw new InvalidOperationException("Font properties were not copied correctly.");
        }

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CopyFontFormatting.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not created.", outputPath);
        }
    }
}
