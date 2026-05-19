using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableCloneExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a paragraph that will serve as the insertion point.
            builder.Writeln("This is the original paragraph.");

            // Build a simple 2x2 table after the paragraph.
            Table originalTable = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();
            builder.EndTable();

            // Clone the existing table (deep clone).
            Table clonedTable = (Table)originalTable.Clone(true);

            // Locate the paragraph after which the cloned table will be inserted.
            Paragraph insertionParagraph = doc.FirstSection.Body.FirstParagraph;

            // Insert the cloned table immediately after the identified paragraph.
            insertionParagraph.ParentNode.InsertAfter(clonedTable, insertionParagraph);

            // Save the resulting document.
            doc.Save("ClonedTable.docx");
        }
    }
}
