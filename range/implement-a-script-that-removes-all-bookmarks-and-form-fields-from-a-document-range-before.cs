using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample content, a bookmark, and a form field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Document start.");

        // Add a bookmark.
        builder.StartBookmark("SampleBookmark");
        builder.Writeln("Text inside the bookmark.");
        builder.EndBookmark("SampleBookmark");

        // Add a checkbox form field.
        builder.InsertCheckBox("CheckBox1", false, 0);
        builder.Writeln("After the form field.");

        // Save the original document (optional, demonstrates the before state).
        doc.Save("OriginalDocument.docx");

        // Remove all bookmarks from the document's range.
        var bookmarkNames = doc.Range.Bookmarks.Select(b => b.Name).ToList();
        foreach (string name in bookmarkNames)
        {
            doc.Range.Bookmarks.Remove(name);
        }

        // Remove all form fields from the document's range.
        var formFields = doc.Range.FormFields.ToArray();
        foreach (var field in formFields)
        {
            field.Remove();
        }

        // Save the cleaned document.
        doc.Save("CleanedDocument.docx");
    }
}
