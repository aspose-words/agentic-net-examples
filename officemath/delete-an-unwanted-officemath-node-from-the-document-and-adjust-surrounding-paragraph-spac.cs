using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class DeleteOfficeMathExample
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph before the first equation.
        builder.Writeln("Paragraph before first equation.");

        // Insert first equation using the deterministic EQ-field bootstrap.
        InsertEquation(builder, @"\f(1,2)");

        // Paragraph between equations.
        builder.Writeln("Paragraph between equations.");

        // Insert second equation using the deterministic EQ-field bootstrap.
        InsertEquation(builder, @"\r(3,x)");

        // Paragraph after the second equation.
        builder.Writeln("Paragraph after second equation.");

        // Save the initial document.
        string originalPath = "Original.docx";
        doc.Save(originalPath, SaveFormat.Docx);

        // Load the document to perform deletion.
        Document loadedDoc = new Document(originalPath);

        // Find all OfficeMath nodes.
        NodeCollection officeMaths = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        if (officeMaths.Count < 2)
            throw new InvalidOperationException("Expected at least two OfficeMath nodes.");

        // Target the second OfficeMath node for deletion.
        OfficeMath secondEquation = (OfficeMath)officeMaths[1];
        Paragraph containingParagraph = secondEquation.ParentNode as Paragraph;

        // Remove the OfficeMath node.
        secondEquation.Remove();

        // Adjust spacing of the surrounding paragraphs.
        if (containingParagraph != null)
        {
            // Increase space after the paragraph that previously contained the equation.
            containingParagraph.ParagraphFormat.SpaceAfter = 12; // 12 points.
        }

        // Optionally adjust the previous paragraph's spacing.
        Paragraph previousParagraph = containingParagraph?.PreviousSibling as Paragraph;
        if (previousParagraph != null)
        {
            previousParagraph.ParagraphFormat.SpaceAfter = 6; // 6 points.
        }

        // Save the modified document.
        string modifiedPath = "Modified.docx";
        loadedDoc.Save(modifiedPath, SaveFormat.Docx);

        // Validate that the output file exists.
        if (!File.Exists(modifiedPath))
            throw new FileNotFoundException("The modified document was not saved.", modifiedPath);
    }

    // Helper method to insert an equation via EQ-field bootstrap.
    private static void InsertEquation(DocumentBuilder builder, string eqSwitch)
    {
        // Insert an EQ field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEq = field as FieldEQ;
        if (fieldEq == null)
            throw new InvalidOperationException("Failed to create FieldEQ.");

        // Write the EQ argument.
        builder.MoveTo(fieldEq.Separator);
        builder.Write(eqSwitch);

        // Update the field so that Aspose.Words can evaluate it.
        field.Update();

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = fieldEq.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            Node fieldStart = fieldEq.Start;
            fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);
            // Remove the original field.
            fieldEq.Remove();
        }
        else
        {
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");
        }
    }
}
