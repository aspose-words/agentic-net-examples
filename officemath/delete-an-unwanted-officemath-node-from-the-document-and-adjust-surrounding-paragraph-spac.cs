using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class DeleteOfficeMathExample
{
    public static void Main()
    {
        // Paths for the documents.
        const string initialPath = "DeletedEquation.docx";
        const string finalPath = "DeletedEquation_Modified.docx";

        // 1. Create a new document and insert two equations using the deterministic EQ-field bootstrap workflow.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with some text.
        builder.Writeln("Paragraph before the first equation.");

        // Insert first equation.
        InsertEquation(builder, @"\f(1,2)"); // Simple fraction 1/2.

        // Second paragraph with some text.
        builder.Writeln("Paragraph between equations.");

        // Insert second equation.
        InsertEquation(builder, @"\r(3,x)"); // Simple root expression.

        // Third paragraph with some text.
        builder.Writeln("Paragraph after the second equation.");

        // Save the initial document.
        doc.Save(initialPath);

        // Validate that the file was created.
        if (!File.Exists(initialPath))
            throw new InvalidOperationException($"Failed to create the output file '{initialPath}'.");

        // 2. Load the document, locate the unwanted OfficeMath node (the second equation), delete it,
        //    and adjust spacing of surrounding paragraphs.
        Document loadedDoc = new Document(initialPath);

        // Get all OfficeMath nodes in the document.
        NodeCollection officeMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);

        // Ensure there are at least two equations to delete the second one.
        if (officeMathNodes.Count < 2)
            throw new InvalidOperationException("The document does not contain the expected number of equations.");

        // Target the second equation (index 1).
        OfficeMath unwantedEquation = (OfficeMath)officeMathNodes[1];

        // Capture the paragraph that contains the unwanted equation before removal.
        Paragraph containingParagraph = unwantedEquation.ParentParagraph;

        // Remove the unwanted OfficeMath node.
        unwantedEquation.Remove();

        // Adjust spacing: add space after the paragraph that previously held the equation.
        if (containingParagraph != null)
        {
            containingParagraph.ParagraphFormat.SpaceAfter = 12f; // 12 points after.
        }

        // Also add space before the next paragraph, if it exists.
        Paragraph nextParagraph = containingParagraph?.NextSibling as Paragraph;
        if (nextParagraph != null)
        {
            nextParagraph.ParagraphFormat.SpaceBefore = 12f; // 12 points before.
        }

        // Save the modified document.
        loadedDoc.Save(finalPath);

        // Validate that the modified file was created.
        if (!File.Exists(finalPath))
            throw new InvalidOperationException($"Failed to create the modified output file '{finalPath}'.");
    }

    // Helper method to insert an equation using the EQ-field bootstrap workflow.
    private static void InsertEquation(DocumentBuilder builder, string eqArgument)
    {
        // Insert an EQ field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        if (field is not FieldEQ fieldEq)
            throw new InvalidOperationException("Failed to create a FieldEQ.");

        // Move to the field separator and write the EQ argument (the equation).
        builder.MoveTo(fieldEq.Separator);
        builder.Write(eqArgument);

        // Update the field so that Aspose.Words can process the EQ code.
        field.Update();

        // Convert the field to a real OfficeMath node.
        OfficeMath? officeMath = fieldEq.AsOfficeMath();

        // If conversion succeeded, insert the OfficeMath node before the field start and remove the field.
        if (officeMath != null)
        {
            Node fieldStart = fieldEq.Start;
            fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);
            fieldEq.Remove();
        }
        else
        {
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");
        }
    }
}
