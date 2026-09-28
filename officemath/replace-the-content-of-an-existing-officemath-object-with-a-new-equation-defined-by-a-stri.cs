using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class OfficeMathReplaceExample
{
    public static void Main()
    {
        const string sourceFile = "Sample.docx";
        const string outputFile = "Output.docx";
        const string bookmarkName = "eq1";
        const string originalEquation = "x+y";
        const string replacementEquation = "a+b=c";

        // 1. Create a sample DOCX with a bookmarked equation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph that contains the bookmarked equation.
        builder.StartBookmark(bookmarkName);
        InsertEquationViaField(builder, originalEquation);
        builder.EndBookmark(bookmarkName);

        // Save the source document.
        doc.Save(sourceFile, SaveFormat.Docx);

        // 2. Load the document and locate the bookmarked equation.
        Document loadedDoc = new Document(sourceFile);

        // Verify the bookmark exists (use LINQ because BookmarkCollection may lack Contains).
        if (!loadedDoc.Range.Bookmarks.Any(b => b.Name == bookmarkName))
            throw new InvalidOperationException($"Bookmark '{bookmarkName}' not found.");

        Bookmark bookmark = loadedDoc.Range.Bookmarks[bookmarkName];

        // Walk up the node tree until we reach the containing paragraph.
        Node node = bookmark.BookmarkStart;
        while (node != null && !(node is Paragraph))
            node = node.ParentNode;
        if (node == null)
            throw new InvalidOperationException("Containing paragraph for the bookmark not found.");

        Paragraph paragraph = (Paragraph)node;

        // Find the top‑level OfficeMath node (MathObjectType == OMathPara) in the paragraph.
        OfficeMath targetMath = null;
        foreach (Node child in paragraph.GetChildNodes(NodeType.OfficeMath, false))
        {
            if (child is OfficeMath om && om.MathObjectType == MathObjectType.OMathPara)
            {
                targetMath = om;
                break;
            }
        }
        if (targetMath == null)
            throw new InvalidOperationException("Target OfficeMath node not found.");

        // 3. Create the replacement OfficeMath via deterministic EQ‑field bootstrap.
        DocumentBuilder replBuilder = new DocumentBuilder(loadedDoc);
        replBuilder.MoveTo(paragraph); // Position at the start of the paragraph.
        replBuilder.InsertField(FieldType.FieldEquation, true);

        // The newly inserted field is the last one in the document.
        Field eqField = loadedDoc.Range.Fields[loadedDoc.Range.Fields.Count - 1];
        if (eqField is not FieldEQ fieldEQ)
            throw new InvalidOperationException("Failed to create FieldEQ.");

        // Write the replacement equation string into the field separator.
        replBuilder.MoveTo(fieldEQ.Separator);
        replBuilder.Write(replacementEquation);

        // Convert the field to a real OfficeMath node.
        OfficeMath newMath = fieldEQ.AsOfficeMath();
        if (newMath == null)
            throw new InvalidOperationException("EQ field conversion returned null.");

        // Insert the new OfficeMath before the old one and clean up.
        ((CompositeNode)paragraph).InsertBefore(newMath, targetMath);
        fieldEQ.Remove();          // Remove the temporary field.
        targetMath.Remove();       // Remove the original equation.

        // 4. Save the modified document.
        loadedDoc.Save(outputFile, SaveFormat.Docx);

        // 5. Reload and verify the replacement.
        Document verifyDoc = new Document(outputFile);

        // Verify the bookmark still exists.
        if (!verifyDoc.Range.Bookmarks.Any(b => b.Name == bookmarkName))
            throw new InvalidOperationException("Bookmark missing after save.");

        Bookmark verifyBookmark = verifyDoc.Range.Bookmarks[bookmarkName];
        Node verifyNode = verifyBookmark.BookmarkStart;
        while (verifyNode != null && !(verifyNode is Paragraph))
            verifyNode = verifyNode.ParentNode;
        if (verifyNode == null)
            throw new InvalidOperationException("Containing paragraph not found after reload.");

        Paragraph verifyParagraph = (Paragraph)verifyNode;
        OfficeMath verifyMath = null;
        foreach (Node child in verifyParagraph.GetChildNodes(NodeType.OfficeMath, false))
        {
            if (child is OfficeMath om && om.MathObjectType == MathObjectType.OMathPara)
            {
                verifyMath = om;
                break;
            }
        }
        if (verifyMath == null)
            throw new InvalidOperationException("Replaced OfficeMath node not found after reload.");

        // Simple validation: ensure the new equation text contains the expected characters.
        string mathText = verifyMath.GetText();
        if (!mathText.Contains("a") || !mathText.Contains("b") || !mathText.Contains("c"))
            throw new InvalidOperationException("Replaced equation does not contain expected content.");

        // All steps completed successfully.
    }

    // Helper that inserts an equation using the deterministic EQ‑field workflow.
    private static void InsertEquationViaField(DocumentBuilder builder, string equation)
    {
        builder.InsertField(FieldType.FieldEquation, true);
        // The inserted field is the last one in the document.
        Field eqField = builder.Document.Range.Fields[builder.Document.Range.Fields.Count - 1];
        if (eqField is not FieldEQ fieldEQ)
            throw new InvalidOperationException("Failed to create FieldEQ.");

        // Write the equation string into the field separator.
        builder.MoveTo(fieldEQ.Separator);
        builder.Write(equation);

        // Convert to OfficeMath and replace the field.
        OfficeMath officeMath = fieldEQ.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field conversion returned null.");

        // Insert the OfficeMath before the field start.
        ((CompositeNode)fieldEQ.Start.ParentNode).InsertBefore(officeMath, fieldEQ.Start);
        // Remove the original field.
        fieldEQ.Remove();
    }
}
