using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableOfContentsForTables
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a TOC field that will list only entries with the "Table" label (table captions).
            // \c "Table" – use the caption label "Table".
            // \h – make entries hyperlinked.
            // \z – hide page numbers in web layout.
            // \u – build the TOC using outline levels.
            builder.InsertTableOfContents("\\c \"Table\" \\h \\z \\u");
            builder.InsertBreak(BreakType.PageBreak);

            // Helper to insert a table with a caption.
            void InsertTableWithCaption(string captionText, int rows, int cols)
            {
                // Insert the caption paragraph using the built‑in Caption style.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Caption;
                builder.Writeln(captionText);

                // Return to the normal style for the table content.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

                // Build the table.
                Table table = builder.StartTable();

                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        builder.InsertCell();
                        builder.Write($"R{r + 1}C{c + 1}");
                    }
                    builder.EndRow();
                }

                builder.EndTable();

                // Add a page break after each table for readability.
                builder.InsertBreak(BreakType.PageBreak);
            }

            // Insert several tables with captions.
            InsertTableWithCaption("Table 1: Sample data set A", 2, 2);
            InsertTableWithCaption("Table 2: Sample data set B", 3, 3);
            InsertTableWithCaption("Table 3: Sample data set C", 2, 4);

            // Update all fields (including the TOC) so the entries are populated.
            doc.UpdateFields();

            // Save the document.
            string outputPath = "TableOfContentsForTables.docx";
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
