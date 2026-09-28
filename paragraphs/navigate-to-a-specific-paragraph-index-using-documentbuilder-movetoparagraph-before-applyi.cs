using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several paragraphs to have content to navigate.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Paragraph {i}");
        }

        // Move the builder to the third paragraph (zero‑based index = 2) at the start of the paragraph.
        builder.MoveToParagraph(2, 0);

        // Apply formatting to the selected paragraph.
        builder.Font.Bold = true;
        builder.Font.Size = 16;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
