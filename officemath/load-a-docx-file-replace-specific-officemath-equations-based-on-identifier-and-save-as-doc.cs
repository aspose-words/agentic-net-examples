using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class OfficeMathReplaceExample
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        const string inputPath = "SampleInput.docx";
        const string outputPath = "SampleOutput.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX containing two identifiable equations.
        // -----------------------------------------------------------------
        CreateSampleDocument(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document, replace the equation with identifier 1 (eq1).
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // Find all OfficeMath nodes.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        foreach (OfficeMath math in mathNodes)
        {
            // Work only with top‑level equations.
            if (math.MathObjectType != MathObjectType.OMathPara)
                continue;

            // Identify the target equation by its CustomNodeId (integer identifier).
            if (math.CustomNodeId == 1)
            {
                // Create a new equation (e.g., a fraction 2/3) in a temporary document.
                OfficeMath newMath = CreateOfficeMath(@"\f(2,3)", 1);

                // Import the new OfficeMath into the target document.
                NodeImporter importer = new NodeImporter(newMath.Document, doc, ImportFormatMode.KeepSourceFormatting);
                OfficeMath importedMath = (OfficeMath)importer.ImportNode(newMath, true);

                // Replace the old equation with the new one.
                math.ParentNode.InsertBefore(importedMath, math);
                math.Remove();

                // No need to continue searching after replacement.
                break;
            }
        }

        // Save the modified document.
        doc.Save(outputPath, SaveFormat.Docx);

        // Simple validation that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }

    // -----------------------------------------------------------------
    // Creates a sample document with two equations identified by
    // integer CustomNodeId values 1 (eq1) and 2 (eq2).
    // -----------------------------------------------------------------
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with an equation identified as 1 (eq1).
        builder.Writeln("First equation:");
        InsertEquation(builder, @"\f(1,4)", 1); // Fraction 1/4

        // Second paragraph with an equation identified as 2 (eq2).
        builder.Writeln();
        builder.Writeln("Second equation:");
        InsertEquation(builder, @"\r(3,x)", 2); // Cube root of x

        // Save the sample document.
        doc.Save(filePath, SaveFormat.Docx);
    }

    // -----------------------------------------------------------------
    // Inserts an EQ field, converts it to a real OfficeMath node,
    // and assigns an integer CustomNodeId for later identification.
    // -----------------------------------------------------------------
    private static void InsertEquation(DocumentBuilder builder, string eqSwitch, int identifier)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ switch arguments.
        builder.MoveTo(field.Separator);
        builder.Write(eqSwitch);

        // Return to the field start position.
        builder.MoveTo(field.Start);

        // Convert the field to OfficeMath.
        OfficeMath officeMath = field.AsOfficeMath();

        if (officeMath != null)
        {
            // Replace the field with the real OfficeMath node.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();

            // Tag the equation for later lookup.
            officeMath.CustomNodeId = identifier;
        }
    }

    // -----------------------------------------------------------------
    // Creates an OfficeMath node in a temporary document using the
    // deterministic EQ‑field bootstrap workflow.
    // -----------------------------------------------------------------
    private static OfficeMath CreateOfficeMath(string eqSwitch, int identifier)
    {
        Document tempDoc = new Document();
        DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);

        // Insert an EQ field.
        FieldEQ field = (FieldEQ)tempBuilder.InsertField(FieldType.FieldEquation, true);
        tempBuilder.MoveTo(field.Separator);
        tempBuilder.Write(eqSwitch);
        tempBuilder.MoveTo(field.Start);

        // Convert to OfficeMath.
        OfficeMath officeMath = field.AsOfficeMath();

        if (officeMath != null)
        {
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();

            // Assign the identifier.
            officeMath.CustomNodeId = identifier;
        }

        return officeMath;
    }
}
