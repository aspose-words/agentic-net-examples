using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Tables;

namespace AsposeWordsCompareTables
{
    class Program
    {
        static void Main()
        {
            // Create temporary file paths.
            string tempDir = Path.GetTempPath();
            string originalPath = Path.Combine(tempDir, "Original.docx");
            string editedPath = Path.Combine(tempDir, "Edited.docx");
            string resultPath = Path.Combine(tempDir, "ComparisonResult.docx");

            // Build the original document with a simple 2x2 table.
            Document docOriginal = new Document();
            DocumentBuilder builderOrig = new DocumentBuilder(docOriginal);
            Table tableOrig = builderOrig.StartTable();
            builderOrig.InsertCell();
            builderOrig.Writeln("A1");
            builderOrig.InsertCell();
            builderOrig.Writeln("B1");
            builderOrig.EndRow();
            builderOrig.InsertCell();
            builderOrig.Writeln("A2");
            builderOrig.InsertCell();
            builderOrig.Writeln("B2");
            builderOrig.EndRow();
            builderOrig.EndTable();
            docOriginal.Save(originalPath);

            // Build the edited document with a 2x3 table (adds an extra column).
            Document docEdited = new Document();
            DocumentBuilder builderEdit = new DocumentBuilder(docEdited);
            Table tableEdit = builderEdit.StartTable();
            builderEdit.InsertCell();
            builderEdit.Writeln("A1");
            builderEdit.InsertCell();
            builderEdit.Writeln("B1");
            builderEdit.InsertCell();
            builderEdit.Writeln("C1");
            builderEdit.EndRow();
            builderEdit.InsertCell();
            builderEdit.Writeln("A2");
            builderEdit.InsertCell();
            builderEdit.Writeln("B2");
            builderEdit.InsertCell();
            builderEdit.Writeln("C2");
            builderEdit.EndRow();
            builderEdit.EndTable();
            docEdited.Save(editedPath);

            // Load the documents from the temporary files.
            Document originalDoc = new Document(originalPath);
            Document editedDoc = new Document(editedPath);

            // Configure comparison options to detect table changes.
            CompareOptions compareOptions = new CompareOptions
            {
                IgnoreTables = false,
                IgnoreFormatting = false,
                IgnoreComments = false,
                IgnoreHeadersAndFooters = false,
                Target = ComparisonTargetType.New
            };

            // Perform the comparison; revisions are added to the original document.
            originalDoc.Compare(editedDoc, "Comparer", DateTime.Now, compareOptions);

            // Count revisions that involve tables.
            int tableRevisions = 0;
            foreach (Revision rev in originalDoc.Revisions)
            {
                if (rev.ParentNode != null && rev.ParentNode.NodeType == NodeType.Table)
                    tableRevisions++;
            }

            // Output results.
            Console.WriteLine($"Total revisions: {originalDoc.Revisions.Count}");
            Console.WriteLine($"Table‑related revisions: {tableRevisions}");

            // Save the comparison result.
            originalDoc.Save(resultPath);
            Console.WriteLine($"Comparison result saved to: {resultPath}");
        }
    }
}
