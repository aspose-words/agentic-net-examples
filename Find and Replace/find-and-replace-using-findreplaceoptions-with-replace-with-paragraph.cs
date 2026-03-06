using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceWithParagraph
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a line that contains a placeholder we will replace.
        builder.Writeln("Dear _Customer_, thank you for your purchase.");

        // Configure find‑replace options.
        FindReplaceOptions options = new FindReplaceOptions();

        // Example: apply paragraph formatting to the new paragraph that will be inserted.
        // Here we center‑align the paragraph that replaces the placeholder.
        options.ApplyParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Replace the placeholder with a name followed by a paragraph break.
        // The meta‑character "&p" inserts a new paragraph.
        doc.Range.Replace("_Customer_", "John Doe&p", options);

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
