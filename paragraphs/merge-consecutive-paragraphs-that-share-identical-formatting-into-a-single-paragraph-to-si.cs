using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a common paragraph style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Font.Size = 12;
        builder.Font.Name = "Arial";
        builder.Font.Color = Color.Black;

        // Two consecutive paragraphs with identical formatting.
        builder.Writeln("Paragraph 1 – same formatting.");
        builder.Writeln("Paragraph 2 – same formatting.");

        // A paragraph with different formatting.
        builder.Font.Color = Color.Red;
        builder.Writeln("Paragraph 3 – different formatting.");

        // Another paragraph that matches the first style but is not consecutive.
        builder.Font.Color = Color.Black;
        builder.Writeln("Paragraph 4 – same as first style, non‑consecutive.");

        // Merge consecutive paragraphs that share identical formatting.
        MergeConsecutiveParagraphs(doc);

        // Save the resulting document.
        doc.Save("MergedParagraphs.docx");
    }

    private static void MergeConsecutiveParagraphs(Document doc)
    {
        // Work with the body of the first section.
        Body body = doc.FirstSection.Body;
        // Iterate while there is a next paragraph to compare.
        for (int i = 0; i < body.Paragraphs.Count - 1; i++)
        {
            Paragraph first = body.Paragraphs[i];
            Paragraph second = body.Paragraphs[i + 1];

            // Compare relevant formatting properties.
            bool sameStyle = first.ParagraphFormat.StyleIdentifier == second.ParagraphFormat.StyleIdentifier;
            bool sameAlignment = first.ParagraphFormat.Alignment == second.ParagraphFormat.Alignment;
            bool sameOutline = first.ParagraphFormat.OutlineLevel == second.ParagraphFormat.OutlineLevel;

            if (sameStyle && sameAlignment && sameOutline)
            {
                // Append all runs from the second paragraph to the first.
                foreach (Run run in second.Runs)
                {
                    // Clone the run to preserve its formatting.
                    first.AppendChild(run.Clone(true));
                }

                // Remove the now‑empty second paragraph.
                second.Remove();

                // After removal, stay at the same index to check the new next paragraph.
                i--;
            }
        }
    }
}
