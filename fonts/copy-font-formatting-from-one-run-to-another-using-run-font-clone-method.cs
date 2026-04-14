using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the first paragraph (created by default).
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create the source run and apply some font formatting.
        Run sourceRun = new Run(doc, "Source text ");
        sourceRun.Font.Name = "Courier New";
        sourceRun.Font.Size = 24;
        sourceRun.Font.Bold = true;
        sourceRun.Font.Color = System.Drawing.Color.Blue; // Fully qualified System.Drawing.Color
        paragraph.AppendChild(sourceRun);

        // Create the destination run with default formatting.
        Run destRun = new Run(doc, "Copied formatting text");
        paragraph.AppendChild(destRun);

        // Copy font properties from the source run to the destination run.
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
            throw new Exception("Font copying failed.");
        }

        // Save the document to a file.
        string outputPath = "CopyFontFormatting.docx";
        doc.Save(outputPath);

        // Verify that the output file exists.
        if (!System.IO.File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
