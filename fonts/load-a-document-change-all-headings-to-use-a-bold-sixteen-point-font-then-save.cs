using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace FontExample
{
    public class Program
    {
        public static void Main()
        {
            // Paths for input and output documents.
            string inputPath = "input.docx";
            string outputPath = "output.docx";

            // Create a sample document if it does not exist.
            if (!File.Exists(inputPath))
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);

                // Add a Heading 1.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
                builder.Writeln("Sample Heading 1");

                // Add a Heading 2.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
                builder.Writeln("Sample Heading 2");

                // Add normal paragraph.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                builder.Writeln("This is a normal paragraph.");

                sampleDoc.Save(inputPath);
            }

            // Load the existing document.
            Document doc = new Document(inputPath);

            // Iterate through all paragraphs and modify headings.
            foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                StyleIdentifier styleId = paragraph.ParagraphFormat.StyleIdentifier;

                // Check if the paragraph is a heading (Heading 1‑9).
                if (styleId >= StyleIdentifier.Heading1 && styleId <= StyleIdentifier.Heading9)
                {
                    // Apply bold and 16‑point size to each run in the heading.
                    foreach (Run run in paragraph.Runs)
                    {
                        run.Font.Bold = true;
                        run.Font.Size = 16;
                    }
                }
            }

            // Save the modified document.
            doc.Save(outputPath);

            // Validate that the output file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine("Document saved successfully to " + outputPath);
            }
            else
            {
                Console.WriteLine("Failed to save the document.");
            }
        }
    }
}
