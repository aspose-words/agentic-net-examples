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

        // Create a paragraph and add it to the document's first section.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create a run with some text.
        Run run = new Run(doc, "Hello World!");

        // Set the run's font color to blue.
        // Aspose.Words.Font.Color expects System.Drawing.Color, so convert from Aspose.Drawing.Color.
        run.Font.Color = System.Drawing.Color.FromArgb(Color.Blue.ToArgb());

        // Append the run to the paragraph.
        paragraph.AppendChild(run);

        // Define the output file path.
        string outputPath = "RunBlueColor.docx";

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + Path.GetFullPath(outputPath));
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }

        // Validate that the font color was set to blue.
        if (run.Font.Color.ToArgb() == System.Drawing.Color.Blue.ToArgb())
        {
            Console.WriteLine("Run font color is correctly set to blue.");
        }
        else
        {
            Console.WriteLine("Run font color is not set to blue.");
        }
    }
}
