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

        // Create the first run with some text.
        Run run1 = new Run(doc, "First run text. ");

        // Obtain the Font object from the first run.
        Aspose.Words.Font sharedFont = run1.Font;

        // Set the desired font size.
        sharedFont.Size = 16;

        // Create an Aspose.Drawing.Color and convert it to System.Drawing.Color.
        Aspose.Drawing.Color aspColor = Aspose.Drawing.Color.Blue;
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());

        // Assign the color to the font.
        sharedFont.Color = sysColor;

        // Create a second run with different text.
        Run run2 = new Run(doc, "Second run text.");

        // Assign the same font properties to the second run for consistency.
        run2.Font.Size = sharedFont.Size;
        run2.Font.Color = sharedFont.Color;

        // Add both runs to the paragraph.
        paragraph.AppendChild(run1);
        paragraph.AppendChild(run2);

        // Add the paragraph to the document body.
        doc.FirstSection.Body.AppendChild(paragraph);

        // Define the output file path.
        string outputPath = "Output.docx";

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (File.Exists(outputPath))
        {
            // Validate that both runs have identical font size and color.
            bool sizeMatch = run1.Font.Size == run2.Font.Size;
            bool colorMatch = run1.Font.Color.ToArgb() == run2.Font.Color.ToArgb();

            Console.WriteLine($"Document saved to '{outputPath}'.");
            Console.WriteLine($"Font size match: {sizeMatch}");
            Console.WriteLine($"Font color match: {colorMatch}");
        }
        else
        {
            Console.WriteLine("Failed to create the output document.");
        }
    }
}
