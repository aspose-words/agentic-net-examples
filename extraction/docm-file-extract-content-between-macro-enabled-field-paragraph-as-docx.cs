using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a macro‑enabled source document in memory.
        Document sourceDoc = new Document();
        sourceDoc.EnsureMinimum();

        // Build the document: insert a MACROBUTTON field followed by a paragraph with some text.
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.InsertField("MACROBUTTON MyMacro Click me");
        builder.Writeln(); // Move to a new paragraph after the field.
        builder.Writeln("This paragraph follows the MACROBUTTON field and will be extracted.");

        // Find the first MACROBUTTON field in the document.
        FieldMacroButton macroField = sourceDoc.Range.Fields
            .FirstOrDefault(f => f.Type == FieldType.FieldMacroButton) as FieldMacroButton;

        if (macroField == null)
        {
            Console.WriteLine("No MACROBUTTON field found in the source document.");
            return;
        }

        // The field end node marks the end of the macro field.
        Node fieldEnd = macroField.End;

        // Determine the paragraph that follows the macro field.
        Paragraph fieldParagraph = fieldEnd.GetAncestor(NodeType.Paragraph) as Paragraph;
        Paragraph nextParagraph = fieldParagraph?.NextSibling as Paragraph;

        if (nextParagraph == null)
        {
            Console.WriteLine("No paragraph found after the MACROBUTTON field.");
            return;
        }

        // Collect all nodes that lie between the field end and the next paragraph (exclusive).
        List<Node> nodesToExtract = new List<Node>();
        Node curNode = fieldEnd.NextSibling;
        while (curNode != null && curNode != nextParagraph)
        {
            nodesToExtract.Add(curNode);
            curNode = curNode.NextSibling;
        }

        // Create a new blank document to hold the extracted content.
        Document extractedDoc = new Document();
        extractedDoc.EnsureMinimum();

        // Import the collected nodes into the new document.
        NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);
        foreach (Node node in nodesToExtract)
        {
            Node importedNode = importer.ImportNode(node, true);
            extractedDoc.FirstSection.Body.AppendChild(importedNode);
        }

        // Save the extracted content as a macro‑free DOCX file.
        extractedDoc.Save("ExtractedContent.docx");

        Console.WriteLine("Extraction complete. Saved as ExtractedContent.docx");
    }
}
