using System;
using Aspose.Words;
using Aspose.Words.Lists;

class ApplyListsToDocument
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document(@"C:\Input\SourceDocument.docx");

        // Create a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Numbered list ----------
        // Add a new numbered list based on the default template.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
        // Apply the list to subsequent paragraphs.
        builder.ListFormat.List = numberedList;

        // Add several items to the numbered list.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Numbered item {i}");
        }

        // End the numbered list.
        builder.ListFormat.RemoveNumbers();

        // Insert a paragraph break between the two lists.
        builder.InsertBreak(BreakType.ParagraphBreak);

        // ---------- Bulleted list ----------
        // Add a new bulleted list based on the default template.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        // Apply the bulleted list.
        builder.ListFormat.List = bulletList;

        // Add several items to the bulleted list.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Bullet item {i}");
        }

        // End the bulleted list.
        builder.ListFormat.RemoveNumbers();

        // Save the modified document.
        doc.Save(@"C:\Output\DocumentWithLists.docx");
    }
}
