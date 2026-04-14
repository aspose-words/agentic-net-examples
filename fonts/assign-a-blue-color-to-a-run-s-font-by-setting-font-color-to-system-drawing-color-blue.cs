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

        // Create a paragraph and add it to the document's body.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create a run with sample text.
        Run run = new Run(doc, "Hello Aspose.Words!");

        // Set the run's font color to blue.
        // Use Aspose.Drawing.Color to create the color, then convert to System.Drawing.Color.
        Aspose.Drawing.Color aspColor = Aspose.Drawing.Color.Blue;
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());
        run.Font.Color = sysColor;

        // Append the run to the paragraph.
        paragraph.AppendChild(run);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output file was not created.");
        }

        // Verify that the font color was set correctly.
        System.Drawing.Color assignedColor = run.Font.Color;
        if (assignedColor.ToArgb() != sysColor.ToArgb())
        {
            throw new Exception("The font color was not set to blue as expected.");
        }
    }
}
