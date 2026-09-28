using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper to insert a simple equation using the EQ-field bootstrap workflow.
        void InsertEquation(string eqSwitch)
        {
            // Insert an empty EQ field.
            Field field = builder.InsertField(FieldType.FieldEquation, true);
            // Cast to FieldEQ to access the separator.
            FieldEQ fieldEq = field as FieldEQ;
            if (fieldEq == null)
                throw new InvalidOperationException("Failed to create FieldEQ.");

            // Write the equation switch into the field separator.
            builder.MoveTo(fieldEq.Separator);
            builder.Write(eqSwitch);

            // Update the field so that Aspose.Words can convert it.
            field.Update();

            // Convert the field to a real OfficeMath node.
            OfficeMath officeMath = fieldEq.AsOfficeMath();
            if (officeMath == null)
                throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

            // Insert the OfficeMath node before the field start.
            fieldEq.Start.ParentNode.InsertBefore(officeMath, fieldEq.Start);
            // Remove the original field.
            fieldEq.Remove();

            // Move the builder to the inserted OfficeMath node and add a new paragraph after it.
            builder.MoveTo(officeMath);
            builder.Writeln();
        }

        // Insert first paragraph with some text.
        builder.Writeln("Paragraph before the first equation.");

        // Insert first equation.
        InsertEquation(@"\f(1,2)"); // Simple fraction 1/2.

        // Insert second paragraph with some text.
        builder.Writeln("Paragraph between equations.");

        // Insert second equation.
        InsertEquation(@"\r(3,x)"); // Simple root expression.

        // Insert a final paragraph.
        builder.Writeln("Paragraph after the second equation.");

        // Save the original document.
        string originalPath = "Original.docx";
        doc.Save(originalPath);

        // Load the document again (optional, we can continue with the same instance).
        Document loadedDoc = new Document(originalPath);

        // Get all top‑level OfficeMath nodes (MathObjectType == OMathPara).
        NodeCollection allMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        var topLevelMathNodes = new System.Collections.Generic.List<OfficeMath>();
        foreach (OfficeMath om in allMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                topLevelMathNodes.Add(om);
        }

        int originalMathCount = topLevelMathNodes.Count;
        if (originalMathCount == 0)
            throw new InvalidOperationException("No top‑level OfficeMath nodes found to delete.");

        // Target the first top‑level OfficeMath node for deletion.
        OfficeMath firstMathNode = topLevelMathNodes[0];
        Paragraph containingParagraph = firstMathNode.GetAncestor(NodeType.Paragraph) as Paragraph;

        // Remove the OfficeMath node.
        firstMathNode.Remove();

        // Adjust spacing of surrounding paragraphs.
        if (containingParagraph != null)
        {
            // Add space after the paragraph that contained the deleted equation.
            containingParagraph.ParagraphFormat.SpaceAfter = 12.0; // points

            // Adjust the previous paragraph, if any.
            Paragraph previousParagraph = containingParagraph.PreviousSibling as Paragraph;
            if (previousParagraph != null)
                previousParagraph.ParagraphFormat.SpaceAfter = 6.0; // points

            // Adjust the next paragraph, if any.
            Paragraph nextParagraph = containingParagraph.NextSibling as Paragraph;
            if (nextParagraph != null)
                nextParagraph.ParagraphFormat.SpaceBefore = 6.0; // points
        }

        // Save the modified document.
        string modifiedPath = "Modified.docx";
        loadedDoc.Save(modifiedPath);

        // Validation.
        if (!File.Exists(modifiedPath))
            throw new FileNotFoundException("Modified document was not saved.", modifiedPath);

        Document validationDoc = new Document(modifiedPath);
        NodeCollection remainingAllMath = validationDoc.GetChildNodes(NodeType.OfficeMath, true);
        int remainingTopLevel = 0;
        foreach (OfficeMath om in remainingAllMath)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                remainingTopLevel++;
        }

        if (remainingTopLevel != originalMathCount - 1)
            throw new InvalidOperationException("The OfficeMath node was not correctly deleted.");

        // Indicate successful completion (no interactive output required).
        Console.WriteLine("OfficeMath node deleted and paragraph spacing adjusted successfully.");
    }
}
