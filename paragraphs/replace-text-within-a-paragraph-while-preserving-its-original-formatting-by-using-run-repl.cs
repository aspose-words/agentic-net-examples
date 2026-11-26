using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ParagraphReplaceExample
{
    class Program
    {
        static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a paragraph with formatted text.
            builder.Font.Size = 16;
            builder.Font.Bold = true;
            builder.Writeln("Hello world! This is a sample paragraph.");

            // Retrieve the first paragraph.
            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

            // Replace the word "Hello" with "Hi" while preserving formatting.
            foreach (Run run in paragraph.Runs)
            {
                if (run.Text.Contains("Hello"))
                {
                    run.Text = run.Text.Replace("Hello", "Hi");
                }
            }

            // Save the modified document.
            doc.Save("Output.docx");
        }
    }
}
