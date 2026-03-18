using System;
using Aspose.Words;
using Aspose.Words.Fields;

class ExtractParagraphAroundFormField
{
    static void Main()
    {
        // Load an existing document that contains form fields.
        Document doc = new Document("Input.docx");

        // Create a DocumentBuilder for navigation and editing.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Choose the form field you want to work with.
        // Here we use the first form field in the collection.
        FormField formField = doc.Range.FormFields[0];

        // Each form field automatically creates a bookmark with the same name as the field.
        string bookmarkName = formField.Name;

        // Move the builder's cursor to the start of that bookmark.
        // The cursor will be positioned just after the bookmark start node.
        builder.MoveToBookmark(bookmarkName);

        // The paragraph that contains the bookmark (and thus the form field) is now the current paragraph.
        Paragraph surroundingParagraph = builder.CurrentParagraph;

        // Extract the full text of the paragraph.
        string paragraphText = surroundingParagraph.GetText();

        // Output the extracted text.
        Console.WriteLine("Paragraph containing the form field:");
        Console.WriteLine(paragraphText.Trim());

        // Optionally, save the document after any modifications.
        doc.Save("Output.docx");
    }
}
