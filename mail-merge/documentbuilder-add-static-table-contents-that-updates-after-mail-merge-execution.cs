using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document and attach a DocumentBuilder to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a static Table of Contents (TOC) field.
        // \o "1-3"  – include heading levels 1 through 3.
        // \h        – make entries hyperlinked.
        // \z \u    – hide page numbers in web layout and use outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add a page break so the TOC appears on its own page.
        builder.InsertBreak(BreakType.PageBreak);

        // -----------------------------------------------------------------
        // Insert headings that will be populated by a mail merge later.
        // The headings must use built‑in heading styles so the TOC can pick them up.
        // -----------------------------------------------------------------

        // First heading – will become a Heading 1 entry in the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.InsertField(" MERGEFIELD Title ");   // placeholder for the title
        builder.Writeln();                           // move to the next line

        // Second heading – will become a Heading 2 entry in the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.InsertField(" MERGEFIELD Subtitle "); // placeholder for the subtitle
        builder.Writeln();

        // Add some regular body text after the headings.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This paragraph will remain unchanged after the mail merge.");

        // -----------------------------------------------------------------
        // Prepare a simple data source for the mail merge.
        // -----------------------------------------------------------------
        DataTable data = new DataTable("MyData");
        data.Columns.Add("Title");
        data.Columns.Add("Subtitle");

        // Add a few records – each record will generate a separate document.
        data.Rows.Add("First Document", "Introduction");
        data.Rows.Add("Second Document", "Details");
        data.Rows.Add("Third Document", "Conclusion");

        // Execute the mail merge. This replaces the MERGEFIELDs with the data above.
        doc.MailMerge.Execute(data);

        // After the merge the headings now contain real text.
        // Update all fields (including the TOC) so the table of contents reflects the new headings.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("MergedWithToc.docx");
    }
}
