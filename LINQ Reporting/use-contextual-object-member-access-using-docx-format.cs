using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bookmark named "MyBookmark".
        builder.StartBookmark("MyBookmark");
        builder.Writeln("This is the bookmarked text.");
        builder.EndBookmark("MyBookmark");

        // Move the cursor to the start of the bookmark using contextual member access.
        builder.MoveToBookmark("MyBookmark");

        // Insert a DOCPROPERTY field at the bookmark location.
        FieldDocProperty docProp = (FieldDocProperty)builder.InsertField(FieldType.FieldDocProperty, true);
        // Set the field result directly (alternatively, set the property name and update).
        docProp.Result = "Sample Property";

        // Insert a page break after the field.
        builder.InsertBreak(BreakType.PageBreak);

        // Add a regular paragraph after the page break.
        builder.Writeln("Content after the bookmark.");

        // Save the document in DOCX format.
        doc.Save("Output.docx");
    }
}
