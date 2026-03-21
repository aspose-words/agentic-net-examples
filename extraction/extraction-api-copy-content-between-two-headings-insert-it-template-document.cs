using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class CopyBetweenHeadings
{
    static void Main()
    {
        // ---------- Create source document ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("Start Heading");          // start heading
        srcBuilder.Writeln("First paragraph.");       // content to copy
        srcBuilder.Writeln("Second paragraph.");      // content to copy
        srcBuilder.Writeln("End Heading");            // end heading

        // ---------- Create template document ----------
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
        tmplBuilder.Writeln("Template start");
        tmplBuilder.StartBookmark("InsertHere");
        tmplBuilder.Writeln("Placeholder (will be replaced)");
        tmplBuilder.EndBookmark("InsertHere");
        tmplBuilder.Writeln("Template end");

        // Define the exact heading texts that mark the start and end of the range to copy.
        const string startHeadingText = "Start Heading";
        const string endHeadingText   = "End Heading";

        // Locate the start and end heading paragraphs in the source document.
        Paragraph startParagraph = FindParagraphByText(srcDoc, startHeadingText);
        Paragraph endParagraph   = FindParagraphByText(srcDoc, endHeadingText);

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Could not find the specified headings.");

        // Prepare a NodeImporter to efficiently import nodes from the source to the template.
        NodeImporter importer = new NodeImporter(srcDoc, templateDoc, ImportFormatMode.KeepSourceFormatting);

        // Move the builder to the bookmark in the template where the content should be placed.
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.MoveToBookmark("InsertHere"); // Ensure the template has a bookmark named "InsertHere".

        // Insert the copied nodes after the bookmark position.
        Node insertionPoint = builder.CurrentParagraph; // Current paragraph is the insertion point.

        // Iterate over all nodes that lie between the two headings (exclusive).
        Node curNode = startParagraph.NextSibling;
        while (curNode != null && curNode != endParagraph)
        {
            // Import the node into the destination document.
            Node importedNode = importer.ImportNode(curNode, true);

            // Insert the imported node after the current insertion point.
            insertionPoint.ParentNode.InsertAfter(importedNode, insertionPoint);
            insertionPoint = importedNode; // Update insertion point for the next node.

            curNode = curNode.NextSibling;
        }

        // Save the modified template document.
        templateDoc.Save("Result.docx");
        Console.WriteLine("Result.docx created successfully.");
    }

    // Helper method to locate a paragraph whose visible text matches the supplied string.
    private static Paragraph FindParagraphByText(Document doc, string text)
    {
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            if (para.GetText().Trim() == text)
                return para;
        }
        return null;
    }
}
