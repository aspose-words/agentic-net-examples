using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class ExtractMixedRange
{
    public static void Main()
    {
        // -------------------- Create a sample source document --------------------
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        // Intro paragraph (outside the extraction range).
        builder.Writeln("Intro paragraph before the table.");

        // Build a 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell A1");
        builder.InsertCell();
        builder.Write("Cell B1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell A2");
        builder.InsertCell();
        builder.Write("Cell B2");
        builder.EndTable();

        // Paragraph that will serve as the end boundary of the extraction.
        builder.Writeln("Target paragraph – this marks the end of the extracted range.");

        // Additional content after the extraction range.
        builder.Writeln("Paragraph after the extracted range.");

        // Save the source document.
        const string sourcePath = "source.docx";
        source.Save(sourcePath);

        // -------------------- Load the source document --------------------
        Document loaded = new Document(sourcePath);

        // Locate the first cell of the first table (start of the range).
        Table firstTable = loaded.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (firstTable == null)
            throw new InvalidOperationException("No table found in the document.");

        Cell startCell = firstTable.FirstRow.FirstCell;
        if (startCell == null)
            throw new InvalidOperationException("The table does not contain any cells.");

        // Locate the paragraph that contains the specific marker text (end of the range).
        Paragraph endParagraph = null;
        foreach (Paragraph para in loaded.FirstSection.Body.Paragraphs)
        {
            if (para.GetText().Contains("Target paragraph"))
            {
                endParagraph = para;
                break;
            }
        }
        if (endParagraph == null)
            throw new InvalidOperationException("End paragraph not found.");

        // -------------------- Prepare the destination document --------------------
        Document result = new Document();
        result.RemoveAllChildren(); // Ensure a clean document.

        // Create a new section and body for the result document.
        Section resultSection = new Section(result);
        result.AppendChild(resultSection);
        Body resultBody = new Body(result);
        resultSection.AppendChild(resultBody);

        // -------------------- Import the required nodes --------------------
        // Use NodeImporter to import nodes from the source into the destination.
        NodeImporter importer = new NodeImporter(loaded, result, ImportFormatMode.KeepSourceFormatting);

        // Import the whole table (which contains the start cell) into the result.
        Node importedTable = importer.ImportNode(firstTable, true);
        resultBody.AppendChild(importedTable);

        // Import the end paragraph into the result.
        Node importedParagraph = importer.ImportNode(endParagraph, true);
        resultBody.AppendChild(importedParagraph);

        // -------------------- Save the extracted content --------------------
        const string resultPath = "extracted.docx";
        result.Save(resultPath);

        // Verify that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");
    }
}
