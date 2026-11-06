using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing; // For Aspose.Drawing.Color

namespace FontColorExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();

            // Create a paragraph and a run with sample text.
            Paragraph paragraph = new Paragraph(doc);
            Run run = new Run(doc, "This text will be blue.");

            // Append the run to the paragraph, and the paragraph to the document body.
            paragraph.AppendChild(run);
            doc.FirstSection.Body.AppendChild(paragraph);

            // Create an Aspose.Drawing.Color representing blue.
            Aspose.Drawing.Color asposeBlue = Aspose.Drawing.Color.Blue;

            // Convert Aspose.Drawing.Color to System.Drawing.Color as required by the Font.Color property.
            System.Drawing.Color systemBlue = System.Drawing.Color.FromArgb(asposeBlue.ToArgb());

            // Assign the blue color to the run's font.
            run.Font.Color = systemBlue;

            // Validate that the color was set correctly.
            if (run.Font.Color.ToArgb() == systemBlue.ToArgb())
            {
                // Validation succeeded.
            }
            else
            {
                // Validation failed; handle as needed.
                throw new InvalidOperationException("Font color assignment failed.");
            }

            // Define output file path.
            string outputPath = "Output.docx";

            // Save the document.
            doc.Save(outputPath);

            // Ensure the output file exists.
            if (File.Exists(outputPath))
            {
                // File saved successfully.
            }
            else
            {
                throw new FileNotFoundException("The document was not saved correctly.", outputPath);
            }
        }
    }
}
