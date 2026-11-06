using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample source document containing various fields.
        // -----------------------------------------------------------------
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        builder.Writeln("Document with fields:");
        builder.InsertField(@"DATE \@ ""MMMM d, yyyy""");
        builder.Writeln();
        builder.InsertField("PAGE");
        builder.Writeln();
        builder.InsertField("MERGEFIELD  MyField");
        builder.Writeln();

        const string sourcePath = "source.docx";
        source.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the source document and evaluate all fields.
        // -----------------------------------------------------------------
        Document loaded = new Document(sourcePath);
        loaded.UpdateFields();

        // -----------------------------------------------------------------
        // 3. Prepare a new empty document that will hold the extracted content.
        // -----------------------------------------------------------------
        Document extracted = new Document();
        extracted.RemoveAllChildren(); // Remove the default empty section.

        // Create a fresh section and body for the destination document.
        Section destSection = new Section(extracted);
        extracted.AppendChild(destSection);
        Body destBody = new Body(extracted);
        destSection.AppendChild(destBody);

        // -----------------------------------------------------------------
        // 4. Clone block‑level nodes (Paragraph, Table, Shape) from the source
        //    document into the destination document, importing them so that
        //    they belong to the destination document.
        // -----------------------------------------------------------------
        NodeCollection sourceNodes = loaded.FirstSection.Body.GetChildNodes(NodeType.Any, true);
        foreach (Node node in sourceNodes)
        {
            // Only block‑level nodes are allowed directly under a Body.
            if (node.NodeType == NodeType.Paragraph ||
                node.NodeType == NodeType.Table ||
                node.NodeType == NodeType.Shape)
            {
                // Import the node into the destination document, preserving formatting.
                Node importedNode = extracted.ImportNode(node, true, ImportFormatMode.KeepSourceFormatting);
                destBody.AppendChild(importedNode);
            }
        }

        // -----------------------------------------------------------------
        // 5. Save the extracted content as a new DOCX file.
        // -----------------------------------------------------------------
        const string outputPath = "extracted.docx";
        extracted.Save(outputPath);

        // -----------------------------------------------------------------
        // 6. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The extracted DOCX file was not created.");

        // Indicate successful completion (no interactive input required).
        Console.WriteLine("Extraction completed successfully.");
    }
}
