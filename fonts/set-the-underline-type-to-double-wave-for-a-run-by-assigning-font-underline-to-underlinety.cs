using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UnderlineDoubleWave.docx");

        // Create a new blank document.
        Aspose.Words.Document doc = new Aspose.Words.Document();

        // Create a run with some text.
        Aspose.Words.Run run = new Aspose.Words.Run(doc, "This run has a double‑wave underline.");

        // Set the underline type to double wave (WavyDouble) for the run's font.
        run.Font.Underline = Aspose.Words.Underline.WavyDouble;

        // Verify that the underline was set correctly.
        if (run.Font.Underline != Aspose.Words.Underline.WavyDouble)
            throw new InvalidOperationException("Failed to set the underline type to double wave.");

        // Append the run to the first paragraph of the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Save the document to the specified path.
        doc.Save(outputPath);

        // Ensure that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
    }
}
