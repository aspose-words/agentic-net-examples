using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // -----------------------------------------------------------------
        // 1. Work with the document's main range.
        // Replace a placeholder string throughout the whole document.
        // -----------------------------------------------------------------
        doc.Range.Replace("(PLACEHOLDER)", "Replaced Text");

        // -----------------------------------------------------------------
        // 2. Format all list items.
        // Iterate over every paragraph, check if it belongs to a list,
        // and then change its font size and color.
        // -----------------------------------------------------------------
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
            {
                // Paragraph itself does not expose a Font property.
                // Apply character formatting to each Run inside the paragraph.
                foreach (Run run in para.Runs)
                {
                    run.Font.Size = 12;
                    run.Font.Color = Color.Blue;
                }
            }
        }

        // -----------------------------------------------------------------
        // 3. Extract the content that lies between two bookmarks.
        // The resulting fragment is saved as a separate document.
        // -----------------------------------------------------------------
        Bookmark startBookmark = doc.Range.Bookmarks["Start"];
        Bookmark endBookmark = doc.Range.Bookmarks["End"];
        if (startBookmark != null && endBookmark != null)
        {
            // The Range of the BookmarkStart node contains everything between the start
            // and the matching BookmarkEnd node (inclusive).
            Document fragment = startBookmark.BookmarkStart.Range.ToDocument();
            fragment.Save("Fragment.docx");
        }

        // -----------------------------------------------------------------
        // 4. Manipulate headers and footers.
        // Replace text in every primary footer of each section.
        // -----------------------------------------------------------------
        foreach (Section section in doc.Sections)
        {
            HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            if (footer != null)
            {
                footer.Range.Replace("Old Footer", "New Footer");
            }
        }

        // -----------------------------------------------------------------
        // 5. Update footnotes.
        // For demonstration, make all footnote text italic.
        // -----------------------------------------------------------------
        foreach (Footnote footnote in doc.GetChildNodes(NodeType.Footnote, true))
        {
            footnote.Font.Italic = true;
        }

        // -----------------------------------------------------------------
        // 6. Extract specific pages while preserving layout, headers,
        // footers, and footnotes.
        // Here we extract the first two pages of the document.
        // -----------------------------------------------------------------
        Document firstTwoPages = doc.ExtractPages(0, 2);
        firstTwoPages.Save("FirstTwoPages.docx");

        // -----------------------------------------------------------------
        // 7. Save the modified original document.
        // -----------------------------------------------------------------
        doc.Save("Output.docx");
    }
}
