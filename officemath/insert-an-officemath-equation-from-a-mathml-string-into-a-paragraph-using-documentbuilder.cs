using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathInsertExample
{
    public static void Main()
    {
        // MathML source string (metadata only – not directly parsed by Aspose.Words)
        // Example MathML: <math><mi>x</mi><mo>=</mo><mfrac><mi>-b</mi><mi>2a</mi></mfrac></math>
        string mathMl = "<math><mi>x</mi><mo>=</mo><mfrac><mi>-b</mi><mi>2a</mi></mfrac></math>";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain the equation.
        builder.Writeln("Paragraph before the equation:");

        // Insert a deterministic EQ field as a bootstrap for OfficeMath.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEq = field as FieldEQ;
        if (fieldEq == null)
            throw new InvalidOperationException("Failed to create FieldEQ.");

        // Move the builder to the field separator and write a simple EQ argument.
        builder.MoveTo(fieldEq.Separator);
        // Simple safe equation that Aspose.Words can convert to OfficeMath.
        builder.Write(@"\f(1,2)");

        // Update the field so that the EQ argument is processed.
        fieldEq.Update();

        // Convert the EQ field to a real OfficeMath node.
        OfficeMath officeMath = fieldEq.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start node.
        Node fieldStart = field.Start;
        fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);

        // Remove the original EQ field, leaving only the real OfficeMath node.
        field.Remove();

        // Save the document.
        string outputPath = "Result.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Validate that the document contains at least one top‑level OfficeMath node.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        if (mathNodes.Count == 0)
            throw new InvalidOperationException("No OfficeMath nodes were found in the saved document.");

        // Output a simple confirmation to the console.
        Console.WriteLine($"Document saved to '{Path.GetFullPath(outputPath)}' with {mathNodes.Count} OfficeMath node(s).");
    }
}
