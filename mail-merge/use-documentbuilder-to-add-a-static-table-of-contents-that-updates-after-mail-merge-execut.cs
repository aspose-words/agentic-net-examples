using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsMailMergeToc
{
    public class Program
    {
        public static void Main()
        {
            // Create a data source for the mail merge.
            DataTable data = new DataTable("MyData");
            data.Columns.Add("Title");
            data.Columns.Add("Content");
            data.Rows.Add("First Chapter", "Content of the first chapter.");
            data.Rows.Add("Second Chapter", "Content of the second chapter.");
            data.Rows.Add("Third Chapter", "Content of the third chapter.");

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a static Table of Contents field.
            // The switches configure the TOC to include headings 1‑3, add hyperlinks, hide page numbers for web view, and use outline levels.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            builder.InsertBreak(BreakType.PageBreak);

            // Begin a mail merge region that will repeat for each row in the data table.
            builder.InsertField(" MERGEFIELD TableStart:MyData");

            // Insert a heading that will appear in the TOC.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.InsertField(" MERGEFIELD Title ");
            builder.Writeln(); // Move to the next line.

            // Insert normal paragraph content.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.InsertField(" MERGEFIELD Content ");
            builder.Writeln();

            // End the mail merge region.
            builder.InsertField(" MERGEFIELD TableEnd:MyData");

            // Execute the mail merge with regions.
            doc.MailMerge.ExecuteWithRegions(data);

            // Update all fields in the document (including the TOC) to reflect the merged content.
            doc.UpdateFields();

            // Save the resulting document.
            doc.Save("MailMergeWithToc.docx");
        }
    }
}
