using Aspose.Words;
using Aspose.Drawing;
using System;
using System.IO;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with a run of text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This run will be blue.");

        // Retrieve the first run in the first paragraph.
        Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

        // Create a blue color using Aspose.Drawing.Color.
        Aspose.Drawing.Color asposeBlue = Aspose.Drawing.Color.Blue;

        // Convert Aspose.Drawing.Color to System.Drawing.Color as required by Font.Color.
        System.Drawing.Color systemBlue = System.Drawing.Color.FromArgb(asposeBlue.ToArgb());

        // Assign the blue color to the run's font.
        run.Font.Color = systemBlue;

        // Validate that the color was set correctly.
        if (run.Font.Color.ToArgb() != systemBlue.ToArgb())
        {
            throw new Exception("Font color assignment failed.");
        }

        // Save the document to a file.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Ensure the output file exists.
        if (!File.Exists(outputPath))
        {
            throw new Exception("Output file was not created.");
        }
    }
}
