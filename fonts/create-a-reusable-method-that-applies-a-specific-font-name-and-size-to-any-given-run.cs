using System;
using System.IO;
using Aspose.Words;

namespace AsposeFontExample
{
    public class Program
    {
        // Reusable method that applies a specific font name and size to a Run.
        public static void ApplyFont(Run run, string fontName, double fontSize)
        {
            // Set font properties using Aspose.Words.Font API.
            run.Font.Name = fontName;
            run.Font.Size = fontSize;

            // Validate that the properties were set correctly.
            if (run.Font.Name != fontName || Math.Abs(run.Font.Size - fontSize) > 0.001)
                throw new InvalidOperationException("Failed to apply font settings to the run.");
        }

        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Ensure the document has at least one paragraph.
            Paragraph para = doc.FirstSection.Body.FirstParagraph ?? new Paragraph(doc);
            if (doc.FirstSection.Body.FirstParagraph == null)
                doc.FirstSection.Body.AppendChild(para);

            // Create a run with sample text.
            Run run = new Run(doc, "Hello Aspose.Words!");
            para.AppendChild(run);

            // Apply the desired font using the reusable method.
            ApplyFont(run, "Courier New", 24);

            // Save the document to a local file.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FormattedRun.docx");
            doc.Save(outputPath);

            // Verify that the output file exists.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The document was not saved correctly.", outputPath);

            // Optional: indicate success (no interactive prompts required).
            Console.WriteLine("Document saved successfully to: " + outputPath);
        }
    }
}
