using System;
using Aspose.Words;
using Aspose.Words.Fields;

class SeeAlsoReferences
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // First section – create a bookmark that we will reference later.
        // -----------------------------------------------------------------
        builder.ParagraphFormat.StyleName = "Heading 1";
        builder.Writeln("Chapter 1: Introduction");

        // Start a bookmark named "Intro".
        builder.StartBookmark("Intro");
        builder.Writeln("This is the introductory chapter. It contains important information.");
        // End the bookmark.
        builder.EndBookmark("Intro");

        // Insert a page break to separate sections.
        builder.InsertBreak(BreakType.PageBreak);

        // -----------------------------------------------------------------
        // Second section – create another bookmark.
        // -----------------------------------------------------------------
        builder.ParagraphFormat.StyleName = "Heading 1";
        builder.Writeln("Chapter 2: Details");

        builder.StartBookmark("Details");
        builder.Writeln("This chapter goes into the details of the topic.");
        builder.EndBookmark("Details");

        // Insert a paragraph that references the first chapter using a REF field.
        builder.Writeln(); // Blank line.
        builder.Write("See also: ");

        // Insert a REF field that points to the "Intro" bookmark.
        FieldRef refField = (FieldRef)builder.InsertField(FieldType.FieldRef, true);
        refField.BookmarkName = "Intro";
        // Optional: make the REF field a hyperlink to the bookmarked text.
        refField.InsertHyperlink = true;

        // Finish the line.
        builder.Writeln();

        // Insert another reference to the second chapter.
        builder.Write("For more details, see: ");
        FieldRef refField2 = (FieldRef)builder.InsertField(FieldType.FieldRef, true);
        refField2.BookmarkName = "Details";
        refField2.InsertHyperlink = true;
        builder.Writeln();

        // -----------------------------------------------------------------
        // Update all fields so that the REF fields display the bookmarked text.
        // -----------------------------------------------------------------
        doc.UpdateFields();

        // Save the document to a file.
        doc.Save("SeeAlsoReferences.docx");
    }
}
