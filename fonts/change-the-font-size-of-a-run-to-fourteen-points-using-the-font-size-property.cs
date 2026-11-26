using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a paragraph to the document.
        Paragraph para = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(para);

        // Create a Run with some text.
        Run run = new Run(doc, "Hello, Aspose.Words!");
        para.AppendChild(run);

        // Change the font size of the Run to fourteen points.
        // Use the fully qualified Aspose.Words.Font type.
        Aspose.Words.Font font = run.Font;
        font.Size = 14.0;

        // Validate that the font size was set correctly.
        if (Math.Abs(font.Size - 14.0) > 0.001)
        {
            throw new InvalidOperationException("Font size was not set to 14 points.");
        }

        // Save the document to a file.
        string outputPath = "OutputDocument.docx";
        doc.Save(outputPath);

        // Ensure the output file exists.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not created.", outputPath);
        }

        // Optionally, indicate success (no user interaction required).
        Console.WriteLine("Document created and saved successfully.");
    }
}
