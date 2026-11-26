using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableCloneExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add an initial paragraph.
            builder.Writeln("Paragraph before the original table.");

            // Build the original table.
            builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Cell 1");
            builder.InsertCell();
            builder.Writeln("Cell 2");
            builder.EndRow();
            builder.InsertCell();
            builder.Writeln("Cell 3");
            builder.InsertCell();
            builder.Writeln("Cell 4");
            builder.EndRow();
            builder.EndTable();

            // Keep a reference to the created table.
            Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Add a paragraph after which the cloned table will be inserted.
            builder.Writeln("Target paragraph. The cloned table will appear after this line.");

            // Locate the target paragraph by its text.
            Paragraph targetParagraph = null;
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                if (para.GetText().Contains("Target paragraph"))
                {
                    targetParagraph = para;
                    break;
                }
            }

            if (targetParagraph == null)
                throw new InvalidOperationException("Target paragraph not found.");

            // Clone the original table (deep clone).
            Node clonedTable = originalTable.Clone(true);

            // Insert the cloned table after the target paragraph.
            // The parent of a paragraph is typically a Body node.
            targetParagraph.ParentNode.InsertAfter(clonedTable, targetParagraph);

            // Save the document.
            string outputPath = "ClonedTable.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not saved.", outputPath);
        }
    }
}
