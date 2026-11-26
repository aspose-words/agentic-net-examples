using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Fields;   // Required for TextFormFieldType and FormFieldCollection

namespace RemoveBookmarksAndFormFields
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and a builder for adding content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph wrapped in a bookmark.
            builder.StartBookmark("SampleBookmark");
            builder.Writeln("This text is inside a bookmark.");
            builder.EndBookmark("SampleBookmark");

            // Insert a text input form field.
            builder.InsertTextInput("SampleFormField", TextFormFieldType.Regular, "", "Default value", 0);
            builder.Writeln(); // Add a line break after the form field.

            // Save the original document (optional, for reference).
            doc.Save("Original.docx");

            // ---------- Remove all bookmarks ----------
            // Collect bookmark names first because removing while iterating changes the collection.
            List<string> bookmarkNames = new List<string>();
            foreach (Bookmark bm in doc.Range.Bookmarks)
                bookmarkNames.Add(bm.Name);

            foreach (string name in bookmarkNames)
                doc.Range.Bookmarks.Remove(name);

            // ---------- Remove all form fields ----------
            FormFieldCollection formFields = doc.Range.FormFields;
            // Remove from the end to avoid index shifting.
            for (int i = formFields.Count - 1; i >= 0; i--)
                formFields.RemoveAt(i);

            // Save the cleaned document.
            doc.Save("Cleaned.docx");
        }
    }
}
