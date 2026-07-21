using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class RemoveBookmarksAndFormFields
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some content with bookmarks.
        builder.StartBookmark("Bookmark1");
        builder.Write("This is the first bookmark.");
        builder.EndBookmark("Bookmark1");
        builder.Writeln();

        builder.StartBookmark("Bookmark2");
        builder.Write("This is the second bookmark.");
        builder.EndBookmark("Bookmark2");
        builder.Writeln();

        // Insert a couple of form fields.
        builder.InsertCheckBox("CheckBox1", false, 50);
        builder.Writeln();
        builder.InsertComboBox("ComboBox1", new[] { "OptionA", "OptionB", "OptionC" }, 0);
        builder.Writeln();

        // At this point the document contains bookmarks and form fields.
        // Remove all bookmarks from the whole document range.
        doc.Range.Bookmarks.Clear();

        // Remove all form fields from the whole document range.
        doc.Range.FormFields.Clear();

        // Save the cleaned document.
        doc.Save("CleanedDocument.docx");
    }
}
