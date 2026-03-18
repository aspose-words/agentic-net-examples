using System;
using System.Drawing;
using Aspose.Words;

class DynamicFontColorExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs to demonstrate the dynamic coloring.
        for (int i = 0; i < 6; i++)
        {
            builder.Writeln($"Paragraph #{i + 1}");
        }

        // Loop through all paragraphs in the first section's body.
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;
        for (int i = 0; i < paragraphs.Count; i++)
        {
            Paragraph para = paragraphs[i];

            // Ensure the paragraph contains at least one run.
            if (para.Runs.Count == 0)
                continue;

            // Choose a color based on the paragraph index.
            // Even index -> DarkBlue, Odd index -> DarkRed.
            Color chosenColor = (i % 2 == 0) ? Color.DarkBlue : Color.DarkRed;

            // Apply the color to every run within the paragraph.
            foreach (Run run in para.Runs)
            {
                run.Font.Color = chosenColor;
            }
        }

        // Save the resulting document.
        doc.Save("DynamicFontColor.docx");
    }
}
