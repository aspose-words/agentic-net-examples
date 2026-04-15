using System;
using System.IO;
using System.Linq;
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

        // Add a paragraph before the first equation.
        builder.Writeln("Paragraph before the first equation.");

        // Insert the first equation (will be kept).
        InsertEquation(builder, @"\f(1,2)"); // Fraction 1/2

        // Add a paragraph between equations.
        builder.Writeln("Paragraph between equations.");

        // Insert the second equation (this one will be removed later).
        InsertEquation(builder, @"\r(3,x)"); // Cube root of x

        // Add a paragraph after the second equation.
        builder.Writeln("Paragraph after the second equation.");

        // Save the initial document (optional, for inspection).
        string initialPath = Path.Combine(Environment.CurrentDirectory, "Initial.docx");
        doc.Save(initialPath);

        // -----------------------------------------------------------------
        // Delete the unwanted OfficeMath node (the second equation) and
        // adjust spacing of surrounding paragraphs.
        // -----------------------------------------------------------------

        // Get only top‑level OfficeMath nodes (MathObjectType.OMathPara).
        var topLevelMath = doc.GetChildNodes(NodeType.OfficeMath, true)
                              .Cast<OfficeMath>()
                              .Where(m => m.MathObjectType == MathObjectType.OMathPara)
                              .ToList();

        int originalCount = topLevelMath.Count;

        if (originalCount == 0)
            throw new InvalidOperationException("No top‑level OfficeMath nodes found to delete.");

        // Identify the unwanted top‑level OfficeMath node (the second one if it exists).
        int indexToDelete = originalCount > 1 ? 1 : 0;
        OfficeMath unwanted = topLevelMath[indexToDelete];

        // Capture surrounding paragraphs before removal.
        Paragraph parentParagraph = unwanted.ParentParagraph;
        Paragraph previousParagraph = parentParagraph?.PreviousSibling as Paragraph;
        Paragraph nextParagraph = parentParagraph?.NextSibling as Paragraph;

        // Remove the OfficeMath node.
        unwanted.Remove();

        // If the parent paragraph became empty, remove it as well.
        if (parentParagraph != null && !parentParagraph.HasChildNodes)
            parentParagraph.Remove();

        // Adjust spacing of surrounding paragraphs.
        if (previousParagraph != null)
            previousParagraph.ParagraphFormat.SpaceAfter = 12f; // 12 points after

        if (nextParagraph != null)
            nextParagraph.ParagraphFormat.SpaceBefore = 12f; // 12 points before

        // Validate that the top‑level OfficeMath count decreased by one.
        int remainingCount = doc.GetChildNodes(NodeType.OfficeMath, true)
                                .Cast<OfficeMath>()
                                .Count(m => m.MathObjectType == MathObjectType.OMathPara);

        if (remainingCount != originalCount - 1)
            throw new InvalidOperationException("Unexpected number of top‑level OfficeMath nodes after deletion.");

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "DeletedOfficeMath.docx");
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }

    // Helper method to insert an OfficeMath equation using the deterministic EQ-field bootstrap workflow.
    private static OfficeMath InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the EQ arguments (field result).
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Ensure the field is up‑to‑date so that AsOfficeMath can parse it.
        field.Update();

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);

        // Remove the original field, leaving only the OfficeMath node.
        field.Remove();

        // Add a paragraph break after the equation for readability.
        builder.Writeln();

        return officeMath;
    }
}
