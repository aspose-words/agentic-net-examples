using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some surrounding text.
        builder.Writeln("Paragraph before the form field.");

        // Insert a text input form field. The field name also creates a bookmark with the same name.
        builder.Write("Enter name: ");
        FormField textField = builder.InsertTextInput(
            "MyName",                     // field name / bookmark name
            TextFormFieldType.Regular,   // field type
            "",                          // format (unused here)
            "John Doe",                  // placeholder text
            50);                         // maximum length

        // Add another paragraph after the field.
        builder.Writeln();
        builder.Writeln("Paragraph after the form field.");

        // Save the document because we have modified it.
        const string filePath = "FormFields.docx";
        doc.Save(filePath);

        // Load the document (optional, we can continue using the same instance).
        Document loadedDoc = new Document(filePath);

        // Navigate to the form field using its automatically created bookmark.
        Bookmark bookmark = loadedDoc.Range.Bookmarks["MyName"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark 'MyName' was not found.");

        // The bookmark start node resides inside the paragraph that contains the form field.
        Node startNode = bookmark.BookmarkStart;
        // Get the paragraph that is an ancestor of the bookmark start node.
        Paragraph containingParagraph = startNode?.GetAncestor(NodeType.Paragraph) as Paragraph;
        if (containingParagraph == null)
            throw new InvalidOperationException("Containing paragraph was not found.");

        // Extract the text of the surrounding paragraph (including the field result).
        string paragraphText = containingParagraph.GetText();

        // Output the extracted paragraph text.
        Console.WriteLine("Extracted paragraph text:");
        Console.WriteLine(paragraphText);
    }
}
