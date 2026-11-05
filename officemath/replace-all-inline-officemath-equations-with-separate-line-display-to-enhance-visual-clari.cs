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
        const string sourceFile = "sample.docx";
        const string outputFile = "output.docx";
        const string targetBookmark = "eq1";
        const string sourceBookmark = "eq2";

        // Clean up any previous files.
        if (File.Exists(sourceFile)) File.Delete(sourceFile);
        if (File.Exists(outputFile)) File.Delete(outputFile);

        // 1. Create a sample DOCX with two equations, each bookmarked.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Equation 1 (bookmark eq1)
        builder.StartBookmark(targetBookmark);
        InsertEquation(builder, "x+y");
        builder.EndBookmark(targetBookmark);

        // Add a blank paragraph between equations.
        builder.Writeln();

        // Equation 2 (bookmark eq2) – will be used as replacement source.
        builder.StartBookmark(sourceBookmark);
        InsertEquation(builder, "p-q");
        builder.EndBookmark(sourceBookmark);

        // Save the source document.
        doc.Save(sourceFile, SaveFormat.Docx);

        // 2. Reload the document.
        Document loadedDoc = new Document(sourceFile);

        // 3. Locate the target bookmark (eq1).
        Bookmark targetBm = loadedDoc.Range.Bookmarks[targetBookmark];
        if (targetBm == null)
            throw new InvalidOperationException($"Bookmark '{targetBookmark}' not found.");

        // Ascend to the containing paragraph.
        Node targetNode = targetBm.BookmarkStart;
        while (targetNode != null && !(targetNode is Paragraph))
            targetNode = targetNode.ParentNode;
        Paragraph targetParagraph = targetNode as Paragraph;
        if (targetParagraph == null)
            throw new InvalidOperationException($"Paragraph for bookmark '{targetBookmark}' not found.");

        // Find the top‑level OfficeMath node inside the paragraph.
        OfficeMath targetMath = targetParagraph.GetChildNodes(NodeType.OfficeMath, false)
            .OfType<OfficeMath>()
            .FirstOrDefault(m => m.MathObjectType == MathObjectType.OMathPara);
        if (targetMath == null)
            throw new InvalidOperationException($"OfficeMath node under bookmark '{targetBookmark}' not found.");

        // 4. Locate the source bookmark (eq2) to clone its equation.
        Bookmark sourceBm = loadedDoc.Range.Bookmarks[sourceBookmark];
        if (sourceBm == null)
            throw new InvalidOperationException($"Bookmark '{sourceBookmark}' not found.");

        Node sourceNode = sourceBm.BookmarkStart;
        while (sourceNode != null && !(sourceNode is Paragraph))
            sourceNode = sourceNode.ParentNode;
        Paragraph sourceParagraph = sourceNode as Paragraph;
        if (sourceParagraph == null)
            throw new InvalidOperationException($"Paragraph for bookmark '{sourceBookmark}' not found.");

        OfficeMath sourceMath = sourceParagraph.GetChildNodes(NodeType.OfficeMath, false)
            .OfType<OfficeMath>()
            .FirstOrDefault(m => m.MathObjectType == MathObjectType.OMathPara);
        if (sourceMath == null)
            throw new InvalidOperationException($"Source OfficeMath node under bookmark '{sourceBookmark}' not found.");

        // Clone the source equation.
        OfficeMath replacementMath = (OfficeMath)sourceMath.Clone(true);
        if (replacementMath == null)
            throw new InvalidOperationException("Failed to clone the source OfficeMath node.");

        // 5. Replace the target equation with the cloned one.
        CompositeNode parent = targetMath.ParentNode as CompositeNode;
        if (parent == null)
            throw new InvalidOperationException("Target OfficeMath does not have a valid parent.");

        parent.InsertBefore(replacementMath, targetMath);
        targetMath.Remove();

        // 6. Save the modified document.
        loadedDoc.Save(outputFile, SaveFormat.Docx);

        // 7. Validate the replacement after reloading.
        Document finalDoc = new Document(outputFile);
        Bookmark finalBm = finalDoc.Range.Bookmarks[targetBookmark];
        if (finalBm == null)
            throw new InvalidOperationException($"Bookmark '{targetBookmark}' missing after save.");

        Node finalNode = finalBm.BookmarkStart;
        while (finalNode != null && !(finalNode is Paragraph))
            finalNode = finalNode.ParentNode;
        Paragraph finalParagraph = finalNode as Paragraph;
        if (finalParagraph == null)
            throw new InvalidOperationException($"Paragraph for bookmark '{targetBookmark}' missing after save.");

        OfficeMath finalMath = finalParagraph.GetChildNodes(NodeType.OfficeMath, false)
            .OfType<OfficeMath>()
            .FirstOrDefault(m => m.MathObjectType == MathObjectType.OMathPara);
        if (finalMath == null)
            throw new InvalidOperationException($"No top‑level OfficeMath found in bookmarked paragraph after replacement.");

        Console.WriteLine("Equation replacement completed successfully.");
    }

    // Helper method to insert an equation using the EQ‑field bootstrap workflow.
    private static void InsertEquation(DocumentBuilder builder, string equationText)
    {
        // Insert an EQ field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        // Move to the field separator and write the equation text.
        builder.MoveTo(field.Separator);
        builder.Write(equationText);
        // Update the field to ensure the equation is parsed.
        field.Update();

        // Convert the field to a real OfficeMath node.
        FieldEQ fieldEq = field as FieldEQ;
        if (fieldEq == null)
            throw new InvalidOperationException("Failed to cast field to FieldEQ.");

        OfficeMath officeMath = fieldEq.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field conversion to OfficeMath returned null.");

        // Insert the OfficeMath node before the field start.
        CompositeNode parent = field.Start.ParentNode as CompositeNode;
        if (parent == null)
            throw new InvalidOperationException("Field start does not have a valid parent.");

        parent.InsertBefore(officeMath, field.Start);
        // Remove the original field.
        field.Remove();
    }
}
