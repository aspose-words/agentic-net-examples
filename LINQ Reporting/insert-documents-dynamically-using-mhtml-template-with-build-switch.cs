using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document and attach a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // MHTML fragments that could come from a template file.
        string mhtmlHeader = "<h1>Report Title</h1>";
        string mhtmlBody   = "<p>This is the body of the report.</p>";
        string mhtmlFooter = $"<p>Generated on {DateTime.Now:yyyy-MM-dd}</p>";

        // Build‑time switches that decide which fragments to insert.
        bool includeHeader = true;
        bool includeBody   = true;
        bool includeFooter = false;

        // Insert the selected fragments using InsertHtml.
        if (includeHeader)
        {
            // Use builder formatting for consistency with surrounding text.
            builder.InsertHtml(mhtmlHeader, HtmlInsertOptions.UseBuilderFormatting);
        }

        if (includeBody)
        {
            // Combine formatting with removal of the extra empty paragraph that InsertHtml adds.
            builder.InsertHtml(mhtmlBody,
                HtmlInsertOptions.UseBuilderFormatting | HtmlInsertOptions.RemoveLastEmptyParagraph);
        }

        if (includeFooter)
        {
            builder.InsertHtml(mhtmlFooter, HtmlInsertOptions.UseBuilderFormatting);
        }

        // Save the assembled document.
        doc.Save("DynamicMhtmlInsert.docx");
    }
}
