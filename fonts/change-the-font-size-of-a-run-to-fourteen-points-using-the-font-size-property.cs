using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a run with some text.
        Run run = new Run(doc, "Hello Aspose.Words!");

        // Change the font size of the run to 14 points.
        run.Font.Size = 14;

        // Append the run to the first paragraph of the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RunFontSize.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            // Load the saved document and confirm the font size.
            Document loadedDoc = new Document(outputPath);
            Run loadedRun = (Run)loadedDoc.GetChild(NodeType.Run, 0, true);
            double fontSize = loadedRun.Font.Size; // Should be 14
        }
    }
}
