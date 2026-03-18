using System;
using Aspose.Words;

class RemoveBookmarksAndFormFields
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Remove all bookmarks from the whole document range.
        // The Clear method removes the bookmarks but leaves the text intact.
        doc.Range.Bookmarks.Clear();

        // Remove all form fields from the whole document range.
        // The Clear method removes the form fields but leaves the field results (text) intact.
        doc.Range.FormFields.Clear();

        // Save the cleaned document.
        doc.Save("Output.docx");
    }
}
