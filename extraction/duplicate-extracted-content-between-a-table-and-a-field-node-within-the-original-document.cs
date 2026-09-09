using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

namespace AsposeWordsExtractionExample
{
    public class Program
    {
        public static void Main()
        {
            // -------------------------------------------------------------
            // 1. Create a sample source document with a table, paragraphs and a field.
            // -------------------------------------------------------------
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            builder.Writeln("Paragraph before the table.");

            // Insert a simple 2x2 table.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("A1");
            builder.InsertCell();
            builder.Write("B1");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("A2");
            builder.InsertCell();
            builder.Write("B2");
            builder.EndRow();
            builder.EndTable();

            // Paragraphs that will be extracted and duplicated.
            builder.Writeln("Middle paragraph 1.");
            builder.Writeln("Middle paragraph 2.");

            // Insert a MERGEFIELD (field node) – this creates its own paragraph.
            builder.InsertField("MERGEFIELD MyField");

            builder.Writeln("Paragraph after the field.");

            // Save the source document (optional, for inspection).
            const string sourcePath = "source.docx";
            sourceDoc.Save(sourcePath);

            // -------------------------------------------------------------
            // 2. Extraction: locate the table and the field's containing paragraph.
            // -------------------------------------------------------------
            // Find the first table in the document.
            Table targetTable = sourceDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
            if (targetTable == null)
                throw new InvalidOperationException("Table not found in the document.");

            // Locate the first field start node and get its ancestor paragraph.
            FieldStart fieldStart = sourceDoc.GetChildNodes(NodeType.FieldStart, true)[0] as FieldStart;
            if (fieldStart == null)
                throw new InvalidOperationException("Field not found in the document.");

            Paragraph fieldParagraph = fieldStart.GetAncestor(NodeType.Paragraph) as Paragraph;
            if (fieldParagraph == null)
                throw new InvalidOperationException("Field's paragraph could not be determined.");

            // -------------------------------------------------------------
            // 3. Collect nodes that lie between the table and the field paragraph.
            // -------------------------------------------------------------
            // Use a forward traversal from the table's next sibling until the field paragraph is reached.
            List<Node> nodesBetween = new List<Node>();
            Node curNode = targetTable.NextSibling;
            while (curNode != null && curNode != fieldParagraph)
            {
                nodesBetween.Add(curNode);
                curNode = curNode.NextSibling;
            }

            if (nodesBetween.Count == 0)
                throw new InvalidOperationException("No nodes found between the table and the field.");

            // -------------------------------------------------------------
            // 4. Duplication: copy the extracted nodes and insert after the field.
            // -------------------------------------------------------------
            // Use NodeImporter to preserve original formatting.
            NodeImporter importer = new NodeImporter(sourceDoc, sourceDoc, ImportFormatMode.KeepSourceFormatting);

            // The parent container where we will insert the duplicated nodes.
            CompositeNode parent = fieldParagraph.ParentNode as CompositeNode;
            if (parent == null)
                throw new InvalidOperationException("Field paragraph does not have a valid parent.");

            // Insert each imported node after the field paragraph, maintaining order.
            Node insertionPoint = fieldParagraph;
            foreach (Node node in nodesBetween)
            {
                Node importedNode = importer.ImportNode(node, true);
                parent.InsertAfter(importedNode, insertionPoint);
                insertionPoint = importedNode;
            }

            // -------------------------------------------------------------
            // 5. Save the modified document.
            // -------------------------------------------------------------
            const string resultPath = "result.docx";
            sourceDoc.Save(resultPath);

            // -------------------------------------------------------------
            // 6. Validation: ensure the duplicated content exists after the field.
            // -------------------------------------------------------------
            if (!File.Exists(resultPath))
                throw new InvalidOperationException("Result document was not created.");

            Document resultDoc = new Document(resultPath);
            Body resultBody = resultDoc.FirstSection.Body;

            // Find the field paragraph again in the result document.
            FieldStart resultFieldStart = resultDoc.GetChildNodes(NodeType.FieldStart, true)[0] as FieldStart;
            Paragraph resultFieldParagraph = resultFieldStart.GetAncestor(NodeType.Paragraph) as Paragraph;
            if (resultFieldParagraph == null)
                throw new InvalidOperationException("Result field paragraph not found.");

            // Use the paragraph collection to locate the duplicated paragraphs.
            ParagraphCollection paragraphs = resultBody.Paragraphs;
            int resultFieldIndex = paragraphs.IndexOf(resultFieldParagraph);
            if (resultFieldIndex < 0 || resultFieldIndex + 2 >= paragraphs.Count)
                throw new InvalidOperationException("Duplicated paragraphs are missing.");

            Paragraph firstDuplicated = paragraphs[resultFieldIndex + 1];
            Paragraph secondDuplicated = paragraphs[resultFieldIndex + 2];

            if (!firstDuplicated.GetText().Contains("Middle paragraph 1") ||
                !secondDuplicated.GetText().Contains("Middle paragraph 2"))
            {
                throw new InvalidOperationException("Duplicated content verification failed.");
            }

            // Program completes without requiring user interaction.
        }
    }
}
