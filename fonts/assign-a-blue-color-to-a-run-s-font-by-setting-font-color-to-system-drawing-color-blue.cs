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

        // Create a paragraph and add it to the document's body.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create a run with some text.
        Run run = new Run(doc, "Hello World!");

        // Create a blue color using Aspose.Drawing.Color.
        Aspose.Drawing.Color asposeBlue = Aspose.Drawing.Color.Blue;

        // Convert Aspose.Drawing.Color to System.Drawing.Color for the Font.Color property.
        System.Drawing.Color systemBlue = System.Drawing.Color.FromArgb(asposeBlue.ToArgb());

        // Assign the blue color to the run's font.
        run.Font.Color = systemBlue;

        // Append the run to the paragraph.
        paragraph.AppendChild(run);

        // Validate that the font color was set correctly.
        if (run.Font.Color.ToArgb() != systemBlue.ToArgb())
        {
            throw new InvalidOperationException("Font color was not set correctly.");
        }

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");

        // Save the document.
        doc.Save(outputPath);

        // Ensure the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not created.", outputPath);
        }
    }
}
