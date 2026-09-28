using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a new paragraph and add it to the document body.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create a run with sample text.
        Run run = new Run(doc, "This text is bold and italic.");

        // Apply bold and italic styles.
        run.Font.Bold = true;
        run.Font.Italic = true;

        // Add the run to the paragraph.
        paragraph.AppendChild(run);

        // Validate that the styles were applied.
        if (run.Font.Bold && run.Font.Italic)
        {
            // Save the document to disk.
            string outputPath = "BoldItalic.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine("Document saved successfully to " + outputPath);
            }
        }
    }
}
