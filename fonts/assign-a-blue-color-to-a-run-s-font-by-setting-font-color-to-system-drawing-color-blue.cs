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

        // Create a run with some text.
        Run run = new Run(doc, "Hello World!");

        // Assign a blue color to the run's font.
        // Font.Color expects a System.Drawing.Color, so convert from Aspose.Drawing.Color.
        run.Font.Color = System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Blue.ToArgb());

        // Append the run to the first paragraph of the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Validate that the color was set correctly.
        System.Drawing.Color expectedColor = System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Blue.ToArgb());
        if (!run.Font.Color.Equals(expectedColor))
        {
            throw new InvalidOperationException("Font color was not set to blue as expected.");
        }

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "RunBlueColor.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
        }
    }
}
