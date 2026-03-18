using System;
using Aspose.Words;
using Aspose.Words.Fields;

class IncludeExternalDocAtBookmark
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a bookmark where the external document will be inserted.
        builder.StartBookmark("InsertionPoint");
        // Placeholder text that will be replaced by the INCLUDE field.
        builder.Write("<<include>>");
        builder.EndBookmark("InsertionPoint");

        // Move the cursor to the start of the bookmark.
        builder.MoveToBookmark("InsertionPoint");

        // Insert an INCLUDE field at the bookmark location.
        FieldInclude includeField = (FieldInclude)builder.InsertField(FieldType.FieldInclude, true);
        // Path to the external DOCX file to be included.
        includeField.SourceFullName = @"C:\Path\To\ExternalDocument.docx";
        // No specific bookmark inside the external document is required; include the whole file.
        includeField.BookmarkName = "";
        // Ensure fields inside the included document are not locked.
        includeField.LockFields = false;
        // Specify the text converter for a Word document.
        includeField.TextConverter = "Microsoft Word";

        // Update all fields so the INCLUDE field resolves.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save(@"C:\Path\To\ResultDocument.docx");
    }
}
