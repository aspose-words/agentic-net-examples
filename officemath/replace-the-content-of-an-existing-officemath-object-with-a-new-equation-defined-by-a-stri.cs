using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class ReplaceOfficeMathExample
{
    public static void Main()
    {
        // Paths for the sample and output documents.
        string samplePath = "Sample.docx";
        string outputPath = "Output.docx";

        // 1. Create a sample DOCX with a bookmarked OfficeMath equation.
        CreateSampleDocument(samplePath);

        // 2. Load the sample document.
        Document doc = new Document(samplePath);

        // 3. Locate the bookmark that identifies the equation to replace.
        const string bookmarkName = "eq1";
        Bookmark bookmark = doc.Range.Bookmarks[bookmarkName];
        if (bookmark == null)
            throw new InvalidOperationException($"Bookmark '{bookmarkName}' not found.");

        // 4. Find the containing paragraph of the bookmark.
        Node node = bookmark.BookmarkStart;
        while (node != null && node.NodeType != NodeType.Paragraph)
            node = node.ParentNode;
        if (node == null)
            throw new InvalidOperationException("Containing paragraph not found.");
        Paragraph paragraph = (Paragraph)node;

        // 5. Locate the top‑level OfficeMath node inside that paragraph.
        OfficeMath targetMath = paragraph.GetChildNodes(NodeType.OfficeMath, false)
                                         .OfType<OfficeMath>()
                                         .FirstOrDefault(m => m.MathObjectType == MathObjectType.OMathPara);
        if (targetMath == null)
            throw new InvalidOperationException("Target OfficeMath node not found.");

        // 6. Create a replacement OfficeMath node by cloning the original.
        //    Cloning is the safest way to obtain a new OfficeMath instance in this workflow.
        OfficeMath replacementMath = (OfficeMath)targetMath.Clone(true);
        if (replacementMath == null)
            throw new InvalidOperationException("Failed to clone OfficeMath.");

        // 7. Insert the replacement before the old node and then remove the old node.
        CompositeNode parent = (CompositeNode)targetMath.ParentNode;
        parent.InsertBefore(replacementMath, targetMath);
        targetMath.Remove();

        // 8. Save the modified document.
        doc.Save(outputPath, SaveFormat.Docx);

        // 9. Reload the saved document and verify the replacement.
        Document resultDoc = new Document(outputPath);
        Bookmark resultBookmark = resultDoc.Range.Bookmarks[bookmarkName];
        if (resultBookmark == null)
            throw new InvalidOperationException("Bookmark missing after save.");

        // Find the paragraph again.
        Node resultNode = resultBookmark.BookmarkStart;
        while (resultNode != null && resultNode.NodeType != NodeType.Paragraph)
            resultNode = resultNode.ParentNode;
        if (resultNode == null)
            throw new InvalidOperationException("Paragraph missing after save.");

        Paragraph resultParagraph = (Paragraph)resultNode;
        OfficeMath finalMath = resultParagraph.GetChildNodes(NodeType.OfficeMath, false)
                                             .OfType<OfficeMath>()
                                             .FirstOrDefault(m => m.MathObjectType == MathObjectType.OMathPara);
        if (finalMath == null)
            throw new InvalidOperationException("Replaced OfficeMath not found after save.");

        // Demonstrate that the OfficeMath node is present.
        Console.WriteLine("Replacement successful. OfficeMath text: " + finalMath.GetText().Trim());
    }

    // Creates a sample document containing a single bookmarked OfficeMath equation.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a bookmark that will surround the equation.
        builder.StartBookmark("eq1");

        // Insert a simple EQ field and convert it to a real OfficeMath node.
        // Use a fraction switch which is known to convert reliably.
        OfficeMath math = CreateOfficeMathFromEq(builder, @"\f(1,2)");
        if (math == null)
            throw new InvalidOperationException("Failed to create initial OfficeMath.");

        // End the bookmark after the equation.
        builder.EndBookmark("eq1");

        // Add a new paragraph so the document is well‑formed.
        builder.Writeln();

        doc.Save(filePath, SaveFormat.Docx);
    }

    // Inserts an EQ field with the specified arguments, converts it to OfficeMath, and returns the OfficeMath node.
    private static OfficeMath CreateOfficeMathFromEq(DocumentBuilder builder, string eqArgs)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArgs);

        // Return the builder to the field start position.
        builder.MoveTo(field.Start);

        // Ensure the field is up‑to‑date so that AsOfficeMath can generate the object.
        field.Update();

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath == null)
            return null;

        // Insert the OfficeMath node before the field and remove the field.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        field.Remove();

        return officeMath;
    }
}
