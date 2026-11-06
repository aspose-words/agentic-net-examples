using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new paragraph.
        builder.Writeln();

        // Create a source run with specific font formatting.
        Run sourceRun = new Run(doc, "Source text");
        sourceRun.Font.Name = "Arial";
        sourceRun.Font.Size = 16;
        sourceRun.Font.Bold = true;
        sourceRun.Font.Italic = true;
        sourceRun.Font.Underline = Underline.Double;

        // Add the source run to the current paragraph.
        builder.CurrentParagraph.AppendChild(sourceRun);

        // Create a destination run with default formatting.
        Run destRun = new Run(doc, "Destination text");
        builder.CurrentParagraph.AppendChild(destRun);

        // Copy font formatting from the source run to the destination run.
        destRun.Font.Name = sourceRun.Font.Name;
        destRun.Font.Size = sourceRun.Font.Size;
        destRun.Font.Bold = sourceRun.Font.Bold;
        destRun.Font.Italic = sourceRun.Font.Italic;
        destRun.Font.Underline = sourceRun.Font.Underline;

        // Validate that the font properties were copied correctly.
        bool fontsMatch =
            destRun.Font.Name == sourceRun.Font.Name &&
            destRun.Font.Size == sourceRun.Font.Size &&
            destRun.Font.Bold == sourceRun.Font.Bold &&
            destRun.Font.Italic == sourceRun.Font.Italic &&
            destRun.Font.Underline == sourceRun.Font.Underline;

        if (fontsMatch)
        {
            // Save the document to disk.
            string outputPath = "CopyFontFormatting.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine("Document saved successfully to " + outputPath);
            }
        }
        else
        {
            Console.WriteLine("Font properties were not copied correctly.");
        }
    }
}
