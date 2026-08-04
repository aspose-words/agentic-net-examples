using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph.
        builder.Writeln("Paragraph before the first equation.");

        // Insert the first (wanted) equation.
        InsertEquation(builder, @"\f(1,2)");

        // Paragraph between equations.
        builder.Writeln("Paragraph between equations.");

        // Insert the second equation which we will delete later.
        InsertEquation(builder, @"\r(2,x)");

        // Paragraph after equations.
        builder.Writeln("Paragraph after equations.");

        // -----------------------------------------------------------------
        // Delete the unwanted OfficeMath node (the second equation).
        // -----------------------------------------------------------------
        NodeCollection officeMaths = doc.GetChildNodes(NodeType.OfficeMath, true);
        if (officeMaths.Count > 1)
        {
            // The second equation is at index 1.
            OfficeMath unwantedMath = (OfficeMath)officeMaths[1];

            // Keep a reference to its parent paragraph before removal.
            Paragraph parentParagraph = unwantedMath.ParentParagraph;

            // Remove the OfficeMath node.
            unwantedMath.Remove();

            // If the paragraph became empty, remove it; otherwise adjust its spacing.
            if (!parentParagraph.HasChildNodes)
            {
                parentParagraph.Remove();
            }
            else
            {
                // Add some space after the paragraph that contained the deleted equation.
                parentParagraph.ParagraphFormat.SpaceAfter = 12; // points
            }

            // Optionally adjust spacing of the preceding paragraph.
            Paragraph previousParagraph = parentParagraph.PreviousSibling as Paragraph;
            if (previousParagraph != null)
            {
                previousParagraph.ParagraphFormat.SpaceAfter = 6; // points
            }
        }

        // Save the resulting document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }

    // Helper method that creates a real OfficeMath node using the deterministic EQ-field bootstrap workflow.
    private static OfficeMath InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        // Replace the field with the real OfficeMath node if conversion succeeded.
        if (officeMath != null)
        {
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }
        else
        {
            // If conversion failed, just remove the field to keep the document clean.
            field.Remove();
        }

        return officeMath;
    }
}
