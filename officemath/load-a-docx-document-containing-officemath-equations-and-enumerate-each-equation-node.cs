using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Path for the sample document (saved in the program's working directory).
        string samplePath = Path.Combine(Directory.GetCurrentDirectory(), "SampleEquations.docx");

        // -----------------------------------------------------------------
        // Create a sample DOCX containing a few OfficeMath equations.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First equation: a simple fraction.
        builder.Writeln("First equation:");
        InsertEquation(builder, @"\f(1,2)");

        // Second equation: a simple root.
        builder.Writeln();
        builder.Writeln("Second equation:");
        InsertEquation(builder, @"\r(3,x)");

        // Save the sample document.
        doc.Save(samplePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the document and enumerate each OfficeMath node.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(samplePath);
        NodeCollection equations = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);

        Console.WriteLine($"Total OfficeMath nodes found: {equations.Count}");
        for (int i = 0; i < equations.Count; i++)
        {
            OfficeMath om = (OfficeMath)equations[i];
            Console.WriteLine($"Equation {i + 1}:");
            Console.WriteLine($"  MathObjectType : {om.MathObjectType}");
            Console.WriteLine($"  DisplayType    : {om.DisplayType}");
            Console.WriteLine($"  Justification  : {om.Justification}");
            Console.WriteLine($"  Text (for display) : {om.GetText()}");
        }
    }

    // Helper method that creates an OfficeMath node using the deterministic EQ-field bootstrap workflow.
    private static void InsertEquation(DocumentBuilder builder, string eqArgument)
    {
        // Insert an empty Equation field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEQ = (FieldEQ)field;

        // Move to the field separator and write the EQ argument.
        builder.MoveTo(fieldEQ.Separator);
        builder.Write(eqArgument);

        // Update the field so that Aspose.Words can convert it to OfficeMath.
        field.Update();

        // Convert the field to a real OfficeMath node.
        OfficeMath? officeMath = fieldEQ.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start and remove the original field.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }
        else
        {
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");
        }
    }
}
