using System;
using Aspose.Words;
using Aspose.Words.Fields;

class SeeAlsoReferencesDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first section with a heading and a bookmark.
        builder.ParagraphFormat.StyleName = "Heading 1";
        builder.StartBookmark("TopicA");
        builder.Writeln("Topic A");
        builder.EndBookmark("TopicA");

        // Add some body text.
        builder.ParagraphFormat.StyleName = "Normal";
        builder.Writeln("This is the main content for Topic A.");

        // Insert a "See also" line that references Topic B and Topic C.
        builder.Writeln(); // blank line
        builder.Write("See also: ");

        // Insert a REF field that points to bookmark "TopicB".
        FieldRef refB = (FieldRef)builder.InsertField(FieldType.FieldRef, true);
        refB.BookmarkName = "TopicB";
        refB.InsertHyperlink = true; // make it clickable
        builder.Write(", ");

        // Insert a REF field that points to bookmark "TopicC".
        FieldRef refC = (FieldRef)builder.InsertField(FieldType.FieldRef, true);
        refC.BookmarkName = "TopicC";
        refC.InsertHyperlink = true;

        // Add a page break before the next topics.
        builder.InsertBreak(BreakType.PageBreak);

        // Insert Topic B.
        builder.ParagraphFormat.StyleName = "Heading 1";
        builder.StartBookmark("TopicB");
        builder.Writeln("Topic B");
        builder.EndBookmark("TopicB");
        builder.ParagraphFormat.StyleName = "Normal";
        builder.Writeln("Details about Topic B.");

        // Insert Topic C.
        builder.InsertBreak(BreakType.PageBreak);
        builder.ParagraphFormat.StyleName = "Heading 1";
        builder.StartBookmark("TopicC");
        builder.Writeln("Topic C");
        builder.EndBookmark("TopicC");
        builder.ParagraphFormat.StyleName = "Normal";
        builder.Writeln("Details about Topic C.");

        // Update all fields so that REF results are calculated.
        doc.UpdateFields();

        // Save the document to a file.
        doc.Save("SeeAlsoReferences.docx");
    }
}
