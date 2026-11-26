using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX with a few equations using the deterministic EQ-field bootstrap workflow.
        string samplePath = "SampleWithEquations.docx";
        CreateSampleDocument(samplePath);

        // Validate that the document was saved.
        if (!File.Exists(samplePath))
            throw new FileNotFoundException($"Failed to create the sample document at '{samplePath}'.");

        // Load the document that contains OfficeMath equations.
        Document doc = new Document(samplePath);

        // Enumerate all OfficeMath nodes in the document.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        Console.WriteLine($"Total OfficeMath nodes found: {officeMathNodes.Count}");

        for (int i = 0; i < officeMathNodes.Count; i++)
        {
            OfficeMath om = officeMathNodes[i] as OfficeMath;
            if (om == null)
                continue;

            // Output basic information about each equation.
            Console.WriteLine($"Equation {i + 1}:");
            Console.WriteLine($"  MathObjectType : {om.MathObjectType}");
            Console.WriteLine($"  DisplayType    : {om.DisplayType}");
            Console.WriteLine($"  Justification  : {om.Justification}");
            Console.WriteLine($"  Text (display) : {om.GetText()}");
        }
    }

    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First equation: \f(1,2)
        InsertEquation(builder, @"\f(1,2)");
        builder.Writeln(); // Add a line break between equations.

        // Second equation: \r(3,x)
        InsertEquation(builder, @"\r(3,x)");
        builder.Writeln();

        // Save the document.
        doc.Save(filePath);
    }

    private static void InsertEquation(DocumentBuilder builder, string eqSwitch)
    {
        // Insert a field of type FieldEquation.
        builder.InsertField(FieldType.FieldEquation, true);

        // Retrieve the inserted field (it is the last field in the document).
        Document doc = builder.Document;
        if (doc.Range.Fields.Count == 0)
            throw new InvalidOperationException("Failed to insert the equation field.");

        Field field = doc.Range.Fields[doc.Range.Fields.Count - 1];

        // Move the builder to the field separator and write the EQ argument.
        builder.MoveTo(field.Separator);
        builder.Write(eqSwitch);

        // Convert the field to a real OfficeMath node.
        FieldEQ fieldEQ = field as FieldEQ;
        OfficeMath officeMath = fieldEQ?.AsOfficeMath();

        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            Node fieldStart = field.Start;
            CompositeNode parent = fieldStart.ParentNode as CompositeNode;
            if (parent == null)
                throw new InvalidOperationException("Field start does not have a composite parent.");

            parent.InsertBefore(officeMath, fieldStart);

            // Remove the original field from the document.
            field.Remove();
        }
        else
        {
            throw new InvalidOperationException("Failed to convert the field to OfficeMath.");
        }
    }
}
