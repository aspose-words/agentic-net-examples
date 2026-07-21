using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class ReplaceOfficeMathExample
{
    public static void Main()
    {
        // Output folder.
        string outputDir = Directory.GetCurrentDirectory();

        // -------------------------------------------------
        // 1. Create a sample document with an initial equation.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Original equation:");

        // Insert an EQ field that will be turned into OfficeMath.
        FieldEQ originalField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Write the EQ argument (a simple fraction 1/2) after the field separator.
        builder.MoveTo(originalField.Separator);
        builder.Write(@"\f(1,2)");
        // Update the field so that Aspose.Words parses the argument.
        originalField.Update();

        // Convert the EQ field to a real OfficeMath node.
        OfficeMath originalMath = originalField.AsOfficeMath();
        if (originalMath == null)
            throw new InvalidOperationException("Failed to convert the original EQ field to OfficeMath.");

        // Replace the field with the OfficeMath node.
        originalField.Start.ParentNode.InsertBefore(originalMath, originalField.Start);
        originalField.Remove();

        // Save the intermediate document (optional).
        string originalPath = Path.Combine(outputDir, "Original.docx");
        doc.Save(originalPath);

        // -------------------------------------------------
        // 2. Replace the content of the existing OfficeMath with a new equation.
        // -------------------------------------------------
        // Locate the top‑level OfficeMath node that we just created.
        OfficeMath targetMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
        if (targetMath == null)
            throw new InvalidOperationException("No OfficeMath node found to replace.");

        // Create a new EQ field that defines the replacement equation.
        builder.MoveToDocumentEnd();
        FieldEQ replacementField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        builder.MoveTo(replacementField.Separator);
        // Example replacement: cube root of x.
        builder.Write(@"\r(3,x)");
        replacementField.Update();

        // Convert the replacement field to OfficeMath.
        OfficeMath replacementMath = replacementField.AsOfficeMath();
        if (replacementMath == null)
            throw new InvalidOperationException("Failed to convert the replacement EQ field to OfficeMath.");

        // Insert the new OfficeMath before the temporary field and remove the field.
        replacementField.Start.ParentNode.InsertBefore(replacementMath, replacementField.Start);
        replacementField.Remove();

        // Replace the original OfficeMath with the new one.
        targetMath.ParentNode.InsertBefore(replacementMath, targetMath);
        targetMath.Remove();

        // -------------------------------------------------
        // 3. Save the final document.
        // -------------------------------------------------
        string resultPath = Path.Combine(outputDir, "Result.docx");
        doc.Save(resultPath);

        Console.WriteLine($"Document created: {resultPath}");
        Console.WriteLine("The original equation has been replaced with the new one.");
    }
}
