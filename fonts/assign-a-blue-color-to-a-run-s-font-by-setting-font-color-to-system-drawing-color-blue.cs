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

        // Get the first paragraph of the document (it always exists in a new document).
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a run with sample text.
        Run run = new Run(doc, "Hello Aspose!");

        // Create Aspose.Drawing.Color.Blue and convert it to System.Drawing.Color.
        Aspose.Drawing.Color asposeBlue = Aspose.Drawing.Color.Blue;
        System.Drawing.Color sysBlue = System.Drawing.Color.FromArgb(asposeBlue.ToArgb());

        // Assign the System.Drawing.Color to the run's font.
        run.Font.Color = sysBlue;

        // Append the run to the paragraph.
        paragraph.AppendChild(run);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RunBlueColor.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
