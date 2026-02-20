using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

namespace AsposeWordsTableSplitExample
{
    // Callback that renames each HTML part file when the document is split.
    class TablePartSavingCallback : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private int _partIndex = 0;

        public TablePartSavingCallback(string baseFileName)
        {
            _baseFileName = baseFileName;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Generate a sequential file name for each part.
            string partFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)}_part{++_partIndex}{Path.GetExtension(args.DocumentPartFileName)}";
            args.DocumentPartFileName = partFileName;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the source HTML document that contains a table.
            Document doc = new Document("Input.html");

            // Get the first table in the document.
            Table originalTable = doc.FirstSection.Body.Tables[0];

            // Split the original table into separate tables – one table per row.
            // The first row stays in the original table; each subsequent row is moved to a new table.
            for (int i = originalTable.Rows.Count - 1; i >= 1; i--)
            {
                // Clone the table structure without its rows.
                Table newTable = (Table)originalTable.Clone(false);
                // Insert the new table after the original table.
                originalTable.ParentNode.InsertAfter(newTable, originalTable);

                // Move the row from the original table to the new table.
                Row rowToMove = originalTable.Rows[i];
                originalTable.Rows.RemoveAt(i);
                newTable.Rows.Add(rowToMove);
            }

            // After each table, insert a section break so that each table can be saved as a separate HTML file.
            DocumentBuilder builder = new DocumentBuilder(doc);
            // Collect tables first because the node list will change while we insert breaks.
            Table[] tables = doc.GetChildNodes(NodeType.Table, true)
                               .OfType<Table>()
                               .ToArray();
            foreach (Table tbl in tables)
            {
                // Move the builder to the last paragraph of the table (the paragraph that follows the table).
                // The table itself does not have a LastParagraph property, so we navigate to the last cell's last paragraph.
                Paragraph lastParagraph = tbl.LastRow?.LastCell?.LastParagraph;
                if (lastParagraph != null)
                {
                    builder.MoveTo(lastParagraph);
                    // Insert a section break (new page) after the table.
                    builder.InsertBreak(BreakType.SectionBreakNewPage);
                }
            }

            // Configure HTML save options to split the document at each section break.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = new TablePartSavingCallback("Output.html")
            };

            // Save the document; Aspose.Words will create multiple HTML files, one per table.
            doc.Save("Output.html", saveOptions);
        }
    }
}
