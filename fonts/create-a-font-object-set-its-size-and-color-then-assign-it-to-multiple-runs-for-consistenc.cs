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

        // Get the first paragraph (created by default).
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create three runs with different text.
        Run run1 = new Run(doc, "Hello ");
        Run run2 = new Run(doc, "World");
        Run run3 = new Run(doc, " Again!");

        // Obtain a Font object from the first run.
        Aspose.Words.Font sharedFont = run1.Font;

        // Set the desired size.
        sharedFont.Size = 24;

        // Create an Aspose.Drawing.Color and convert it to System.Drawing.Color.
        Aspose.Drawing.Color aspColor = Aspose.Drawing.Color.Blue;
        sharedFont.Color = System.Drawing.Color.FromArgb(aspColor.ToArgb());

        // Apply the same font settings to the other runs.
        run2.Font.Size = sharedFont.Size;
        run2.Font.Color = sharedFont.Color;

        run3.Font.Size = sharedFont.Size;
        run3.Font.Color = sharedFont.Color;

        // Append the runs to the paragraph.
        paragraph.AppendChild(run1);
        paragraph.AppendChild(run2);
        paragraph.AppendChild(run3);

        // Save the document.
        string outputPath = "FontRuns.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
    }
}
