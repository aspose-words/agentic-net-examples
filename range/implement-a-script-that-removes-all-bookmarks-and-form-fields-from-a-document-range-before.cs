using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing bookmarks and form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");

        // Insert a bookmark.
        builder.StartBookmark("SampleBookmark");
        builder.Writeln("Text inside the bookmark.");
        builder.EndBookmark("SampleBookmark");

        // Insert a text input form field (maxLength set to 0 for unlimited length).
        builder.InsertTextInput("TextInputField", TextFormFieldType.Regular, "", "Default value", 0);

        // Insert a checkbox form field.
        builder.InsertCheckBox("CheckBoxField", true, 0);

        // Save the original document.
        string originalPath = "Original.docx";
        doc.Save(originalPath);

        // Load the document for cleanup.
        Document cleanDoc = new Document(originalPath);

        // Remove all bookmarks from the document range.
        List<string> bookmarkNames = new List<string>();
        foreach (Bookmark bm in cleanDoc.Range.Bookmarks)
        {
            bookmarkNames.Add(bm.Name);
        }
        foreach (string name in bookmarkNames)
        {
            cleanDoc.Range.Bookmarks.Remove(name);
        }

        // Remove all form fields from the document range.
        FormFieldCollection formFields = cleanDoc.Range.FormFields;
        for (int i = formFields.Count - 1; i >= 0; i--)
        {
            formFields.RemoveAt(i);
        }

        // Save the cleaned document.
        string cleanedPath = "Cleaned.docx";
        cleanDoc.Save(cleanedPath);

        // Indicate completion.
        Console.WriteLine($"Cleaned document saved to '{cleanedPath}'.");
    }
}
