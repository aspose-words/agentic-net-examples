using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class ImportMultipleTables
{
    public static void Main()
    {
        // Define file names for the sample source documents and the merged result.
        string sourcePath1 = Path.Combine(Directory.GetCurrentDirectory(), "Source1.docx");
        string sourcePath2 = Path.Combine(Directory.GetCurrentDirectory(), "Source2.docx");
        string mergedPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedTables.docx");

        // -----------------------------------------------------------------
        // 1. Create two sample documents, each containing a single table.
        // -----------------------------------------------------------------
        CreateSampleDocument(sourcePath1, "First", "Table from Document 1");
        CreateSampleDocument(sourcePath2, "Second", "Table from Document 2");

        // -----------------------------------------------------------------
        // 2. Load the source documents.
        // -----------------------------------------------------------------
        Document srcDoc1 = new Document(sourcePath1);
        Document srcDoc2 = new Document(sourcePath2);

        // -----------------------------------------------------------------
        // 3. Create the destination document that will hold the merged tables.
        // -----------------------------------------------------------------
        Document dstDoc = new Document();

        // -----------------------------------------------------------------
        // 4. Import the first table from each source document into the destination.
        // -----------------------------------------------------------------
        ImportTableIntoDocument(srcDoc1, dstDoc);
        ImportTableIntoDocument(srcDoc2, dstDoc);

        // -----------------------------------------------------------------
        // 5. Save the merged document.
        // -----------------------------------------------------------------
        dstDoc.Save(mergedPath);

        // -----------------------------------------------------------------
        // 6. Verify that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("The merged document was not saved correctly.");
    }

    // Creates a simple document with a single table containing the supplied text.
    private static void CreateSampleDocument(string filePath, string headerText, string cellText)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading so each source document is distinguishable.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln(headerText);

        // Build a 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write(cellText);
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();
        builder.EndTable();

        // Save the sample document.
        doc.Save(filePath);
    }

    // Imports the first table from a source document into the destination document,
    // preserving the original formatting.
    private static void ImportTableIntoDocument(Document srcDoc, Document dstDoc)
    {
        // Locate the first table in the source document.
        Table srcTable = srcDoc.FirstSection.Body.Tables[0];
        if (srcTable == null)
            throw new InvalidOperationException("Source document does not contain a table.");

        // Use NodeImporter to import the table node with source formatting.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedNode = importer.ImportNode(srcTable, true);

        // Append the imported table to the end of the destination document's body.
        dstDoc.FirstSection.Body.AppendChild(importedNode);
    }
}
