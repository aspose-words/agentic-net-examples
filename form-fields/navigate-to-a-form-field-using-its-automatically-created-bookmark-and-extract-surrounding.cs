using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph before the form field.
        builder.Writeln("This is the first paragraph.");

        // Insert a text input form field. This also creates a bookmark with the same name.
        string fieldName = "MyTextField";
        builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", "Default value", 0);

        // Add a paragraph after the form field.
        builder.Writeln("This is the paragraph after the form field.");

        // Save the document.
        doc.Save("FormFieldDemo.docx");

        // Access the form field through the FormFields collection.
        FormField formField = doc.Range.FormFields[fieldName];
        if (formField == null)
        {
            throw new InvalidOperationException($"Form field '{fieldName}' was not found.");
        }

        // The form field automatically creates a bookmark with the same name.
        // Use the bookmark to locate the paragraph that contains the field.
        Bookmark bookmark = doc.Range.Bookmarks[fieldName];
        if (bookmark == null)
        {
            throw new InvalidOperationException($"Bookmark for form field '{fieldName}' was not found.");
        }

        // The start node of the bookmark is a Run; its parent is the containing Paragraph.
        Node startNode = bookmark.BookmarkStart;
        if (startNode == null || startNode.ParentNode == null)
        {
            throw new InvalidOperationException("Unable to locate the paragraph containing the form field.");
        }

        Paragraph containingParagraph = startNode.ParentNode as Paragraph;
        if (containingParagraph == null)
        {
            throw new InvalidOperationException("The parent node is not a paragraph.");
        }

        // Extract the text of the surrounding paragraph.
        string paragraphText = containingParagraph.GetText();

        // Output the extracted paragraph text.
        Console.WriteLine("Paragraph containing the form field:");
        Console.WriteLine(paragraphText);
    }
}
