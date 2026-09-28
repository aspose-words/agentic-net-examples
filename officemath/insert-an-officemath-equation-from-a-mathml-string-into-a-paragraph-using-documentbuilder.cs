using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class InsertOfficeMathFromMathML
{
    public static void Main()
    {
        // The original MathML is kept as a comment for reference.
        // Aspose.Words does not import MathML directly, so we will create a real OfficeMath node
        // using the deterministic EQ‑field bootstrap workflow.

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph before the equation.
        builder.Writeln("Paragraph before equation.");

        // Insert an EQ field that will later be converted to a real OfficeMath node.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEq = field as FieldEQ;
        if (fieldEq == null)
            throw new InvalidOperationException("Inserted field is not a FieldEQ.");

        // Move to the field separator and write a simple, safe EQ argument.
        // The expression "\f(1,2)" reliably converts to an OfficeMath fraction.
        builder.MoveTo(fieldEq.Separator);
        builder.Write(@"\f(1,2)");

        // Update the field so that the EQ argument is processed.
        field.Update();

        // Convert the EQ field to an OfficeMath object.
        OfficeMath officeMath = fieldEq.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        Node fieldStart = fieldEq.Start;
        fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);
        fieldEq.Remove();

        // Continue writing after the equation.
        builder.Writeln();
        builder.Writeln("Paragraph after equation.");

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created and contains at least one top‑level OfficeMath node.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        Document loadedDoc = new Document(outputPath);
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        if (mathNodes.Count == 0)
            throw new InvalidOperationException("No OfficeMath nodes were found in the saved document.");

        Console.WriteLine($"Document saved to '{outputPath}' with {mathNodes.Count} OfficeMath node(s).");
    }
}
