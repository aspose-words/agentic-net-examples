using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

namespace AsposeWordsRangeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Section 1 content.
            builder.Writeln("Section 1 - First line.");
            builder.Writeln("Section 1 - Second line.");

            // Insert a continuous section break to start Section 2.
            builder.InsertBreak(BreakType.SectionBreakContinuous);

            // Section 2 content.
            builder.Writeln("Section 2 - First line.");
            builder.Writeln("Section 2 - Second line.");

            // Insert another continuous section break to start Section 3.
            builder.InsertBreak(BreakType.SectionBreakContinuous);

            // Section 3 content.
            builder.Writeln("Section 3 - Only line.");

            // Extract plain‑text from each section's Range.
            StringBuilder plainTextBuilder = new StringBuilder();

            for (int i = 0; i < doc.Sections.Count; i++)
            {
                // Get the text of the current section.
                string sectionText = doc.Sections[i].Range.Text;

                // Append a header for clarity (optional).
                plainTextBuilder.AppendLine($"--- Section {i + 1} ---");
                plainTextBuilder.AppendLine(sectionText.Trim());
                plainTextBuilder.AppendLine(); // Add an empty line between sections.
            }

            // Define output file path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "PlainTextOutput.txt");

            // Write the combined plain‑text to the file.
            File.WriteAllText(outputPath, plainTextBuilder.ToString());

            // Optionally, display the result path.
            Console.WriteLine($"Plain‑text document saved to: {outputPath}");
        }
    }
}
