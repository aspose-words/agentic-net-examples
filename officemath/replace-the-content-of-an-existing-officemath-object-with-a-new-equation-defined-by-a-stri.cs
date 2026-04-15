using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class ReplaceOfficeMathExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Step 1: Insert an initial equation (fraction 1/2) using the deterministic EQ‑field bootstrap workflow.
        // -------------------------------------------------
        FieldEQ initialField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Write the EQ switch and its arguments.
        builder.MoveTo(initialField.Separator);
        builder.Write(@"\f(1,2)"); // Fraction 1 over 2.
        // Ensure the field is up‑to‑date before conversion.
        initialField.Update();

        // Convert the EQ field to a real OfficeMath node.
        OfficeMath initialMath = initialField.AsOfficeMath();
        if (initialMath == null)
            throw new InvalidOperationException("Failed to create the initial OfficeMath node.");

        // Replace the field with the real OfficeMath node.
        initialField.Start.ParentNode.InsertBefore(initialMath, initialField.Start);
        initialField.Remove();

        // -------------------------------------------------
        // Step 2: Replace the existing OfficeMath with a new equation defined by a string.
        // -------------------------------------------------
        // New equation string (cube root of x) – stored as metadata.
        string newEquation = @"\r(3,x)";

        // Locate the existing top‑level OfficeMath node.
        OfficeMath existingMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
        if (existingMath == null)
            throw new InvalidOperationException("No OfficeMath node found to replace.");

        // Move the builder to the paragraph that contains the existing OfficeMath.
        builder.MoveTo(existingMath);

        // Insert a new EQ field for the replacement equation.
        FieldEQ replacementField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        builder.MoveTo(replacementField.Separator);
        builder.Write(newEquation);
        replacementField.Update();

        // Convert the new EQ field to a real OfficeMath node.
        OfficeMath replacementMath = replacementField.AsOfficeMath();
        if (replacementMath == null)
            throw new InvalidOperationException("Failed to create the replacement OfficeMath node.");

        // Insert the new OfficeMath before the old one, then remove the old node and the temporary field.
        existingMath.ParentNode.InsertBefore(replacementMath, existingMath);
        existingMath.Remove();
        replacementField.Remove();

        // -------------------------------------------------
        // Step 3: Save the document.
        // -------------------------------------------------
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ReplaceOfficeMath.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
