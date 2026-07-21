using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the sample files in the current directory.
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.doc");
        string docxPath = Path.Combine(Directory.GetCurrentDirectory(), "Revised.docx");
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");

        // ---------- Create the original DOC file with a 2x2 table ----------
        Document originalDoc = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(originalDoc);
        builder1.Writeln("Original document with a table:");
        builder1.StartTable();
        builder1.InsertCell();
        builder1.Write("Cell 1A");
        builder1.InsertCell();
        builder1.Write("Cell 1B");
        builder1.EndRow(); // First row finished
        builder1.InsertCell();
        builder1.Write("Cell 2A");
        builder1.InsertCell();
        builder1.Write("Cell 2B");
        builder1.EndRow(); // Second row finished
        builder1.EndTable(); // Close the table
        originalDoc.Save(docPath, SaveFormat.Doc);

        // ---------- Create the revised DOCX file with a modified table ----------
        Document revisedDoc = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revisedDoc);
        builder2.Writeln("Revised document with a changed table:");
        builder2.StartTable();
        builder2.InsertCell();
        builder2.Write("Cell 1A - edited"); // changed text
        builder2.InsertCell();
        builder2.Write("Cell 1B");
        builder2.EndRow(); // First row finished
        builder2.InsertCell();
        builder2.Write("Cell 2A");
        builder2.InsertCell();
        builder2.Write("Cell 2B");
        builder2.EndRow(); // Second row finished
        // Add an extra row to create a structural difference.
        builder2.InsertCell();
        builder2.Write("Cell 3A");
        builder2.InsertCell();
        builder2.Write("Cell 3B");
        builder2.EndRow(); // Third row finished
        builder2.EndTable(); // Close the table
        revisedDoc.Save(docxPath, SaveFormat.Docx);

        // ---------- Load the two documents ----------
        Document docToCompare = new Document(docPath);
        Document docRevised = new Document(docxPath);

        // ---------- Perform the comparison ----------
        // The original document will receive revisions describing the differences.
        docToCompare.Compare(docRevised, "Comparer", DateTime.Now);

        // ---------- Verify that at least one revision exists ----------
        if (docToCompare.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison, but none were found.");

        // ---------- Verify that a table‑related revision was detected ----------
        bool tableRevisionFound = false;
        foreach (Revision rev in docToCompare.Revisions)
        {
            // Table structure changes can appear on Table, Row or Cell nodes.
            if (rev.ParentNode != null)
            {
                NodeType type = rev.ParentNode.NodeType;
                if (type == NodeType.Table || type == NodeType.Row || type == NodeType.Cell)
                {
                    tableRevisionFound = true;
                    break;
                }
            }
        }

        if (!tableRevisionFound)
            throw new InvalidOperationException("Table structure differences were not detected as revisions.");

        // ---------- Save the comparison result ----------
        docToCompare.Save(resultPath, SaveFormat.Docx);

        Console.WriteLine("Comparison completed successfully. Revisions detected and result saved to:");
        Console.WriteLine(resultPath);
    }
}
