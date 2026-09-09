using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text with a bookmark.
        builder.Write("Hello ");
        builder.StartBookmark("MyBookmark");
        builder.Write("World");
        builder.EndBookmark("MyBookmark");
        builder.Writeln();

        // Insert a check box form field.
        builder.InsertCheckBox("CheckBox1", false, 50);
        builder.Writeln();

        // At this point the document contains a bookmark and a form field.
        // Remove all bookmarks from the whole document range.
        doc.Range.Bookmarks.Clear();

        // Remove all form fields from the whole document range.
        doc.Range.FormFields.Clear();

        // Save the cleaned document.
        doc.Save("Output.docx");
    }
}
