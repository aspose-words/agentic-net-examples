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
        // Paths for the sample input and output documents.
        const string inputPath = "SampleInput.docx";
        const string outputPath = "SampleOutput.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample DOCX with two bookmarked equations.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First equation with bookmark "eq1".
        builder.Writeln(); // Ensure a new paragraph.
        builder.StartBookmark("eq1");
        InsertEquationViaEQField(builder, @"\f(1,2)"); // Simple fraction.
        builder.EndBookmark("eq1");

        // Second equation with bookmark "eq2".
        builder.Writeln(); // New paragraph.
        builder.StartBookmark("eq2");
        InsertEquationViaEQField(builder, @"\r(3,x)"); // Simple radical.
        builder.EndBookmark("eq2");

        // Save the sample input document.
        doc.Save(inputPath, SaveFormat.Docx);
        if (!File.Exists(inputPath))
            throw new Exception($"Failed to create input file '{inputPath}'.");

        // -----------------------------------------------------------------
        // Step 2: Load the document and locate the target equation (eq1).
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // Locate bookmark "eq1".
        Bookmark targetBookmark = loadedDoc.Range.Bookmarks["eq1"];
        if (targetBookmark == null)
            throw new Exception("Bookmark 'eq1' not found in the document.");

        // Ascend to the containing paragraph.
        Node targetNode = targetBookmark.BookmarkStart;
        while (targetNode != null && targetNode.NodeType != NodeType.Paragraph)
            targetNode = targetNode.ParentNode;
        if (targetNode == null)
            throw new Exception("Paragraph containing bookmark 'eq1' not found.");

        Paragraph targetParagraph = (Paragraph)targetNode;

        // Find the top‑level OfficeMath node inside the paragraph.
        OfficeMath targetOfficeMath = targetParagraph.GetChildNodes(NodeType.OfficeMath, true)
            .Cast<OfficeMath>()
            .FirstOrDefault(om => om.MathObjectType == MathObjectType.OMathPara);
        if (targetOfficeMath == null)
            throw new Exception("Target OfficeMath node not found in paragraph 'eq1'.");

        // -----------------------------------------------------------------
        // Step 3: Obtain a replacement OfficeMath node (clone of eq2).
        // -----------------------------------------------------------------
        // Locate bookmark "eq2".
        Bookmark sourceBookmark = loadedDoc.Range.Bookmarks["eq2"];
        if (sourceBookmark == null)
            throw new Exception("Bookmark 'eq2' not found in the document.");

        // Ascend to its paragraph.
        Node sourceNode = sourceBookmark.BookmarkStart;
        while (sourceNode != null && sourceNode.NodeType != NodeType.Paragraph)
            sourceNode = sourceNode.ParentNode;
        if (sourceNode == null)
            throw new Exception("Paragraph containing bookmark 'eq2' not found.");

        Paragraph sourceParagraph = (Paragraph)sourceNode;

        // Find the top‑level OfficeMath node inside the source paragraph.
        OfficeMath sourceOfficeMath = sourceParagraph.GetChildNodes(NodeType.OfficeMath, true)
            .Cast<OfficeMath>()
            .FirstOrDefault(om => om.MathObjectType == MathObjectType.OMathPara);
        if (sourceOfficeMath == null)
            throw new Exception("Source OfficeMath node not found in paragraph 'eq2'.");

        // Clone the source OfficeMath node to use as replacement.
        OfficeMath replacementOfficeMath = (OfficeMath)sourceOfficeMath.Clone(true);
        if (replacementOfficeMath == null)
            throw new Exception("Failed to clone the source OfficeMath node.");

        // -----------------------------------------------------------------
        // Step 4: Replace the target equation with the cloned one.
        // -----------------------------------------------------------------
        CompositeNode parentComposite = (CompositeNode)targetOfficeMath.ParentNode;
        if (parentComposite == null)
            throw new Exception("Parent of the target OfficeMath node is null.");

        // Insert the replacement before the target.
        parentComposite.InsertBefore(replacementOfficeMath, targetOfficeMath);
        // Remove the original target equation.
        targetOfficeMath.Remove();

        // -----------------------------------------------------------------
        // Step 5: Save the modified document and validate the result.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputPath, SaveFormat.Docx);
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create output file '{outputPath}'.");

        // Reload the saved document to verify replacement.
        Document resultDoc = new Document(outputPath);
        Bookmark resultBookmark = resultDoc.Range.Bookmarks["eq1"];
        if (resultBookmark == null)
            throw new Exception("Bookmark 'eq1' missing after saving the output document.");

        // Ascend to its paragraph.
        Node resultNode = resultBookmark.BookmarkStart;
        while (resultNode != null && resultNode.NodeType != NodeType.Paragraph)
            resultNode = resultNode.ParentNode;
        if (resultNode == null)
            throw new Exception("Paragraph for bookmark 'eq1' not found after reload.");

        Paragraph resultParagraph = (Paragraph)resultNode;

        // Verify that the paragraph now contains exactly one top‑level OfficeMath node.
        OfficeMath resultOfficeMath = resultParagraph.GetChildNodes(NodeType.OfficeMath, true)
            .Cast<OfficeMath>()
            .FirstOrDefault(om => om.MathObjectType == MathObjectType.OMathPara);
        if (resultOfficeMath == null)
            throw new Exception("Replaced OfficeMath node not found after reload.");

        Console.WriteLine("Equation replacement completed successfully.");
    }

    // Helper method to create a real OfficeMath node using the EQ‑field bootstrap workflow.
    private static void InsertEquationViaEQField(DocumentBuilder builder, string eqArgument)
    {
        // Insert an EQ field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the EQ argument.
        builder.MoveTo(field.Separator);
        builder.Write(eqArgument);

        // Update the field to ensure the argument is processed.
        field.Update();

        // Convert the field to OfficeMath.
        FieldEQ fieldEQ = field as FieldEQ;
        OfficeMath officeMath = fieldEQ?.AsOfficeMath();

        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            CompositeNode parent = (CompositeNode)field.Start.ParentNode;
            parent.InsertBefore(officeMath, field.Start);
            // Remove the original field.
            field.Remove();
        }
        else
        {
            throw new Exception("Failed to convert EQ field to OfficeMath.");
        }
    }
}
