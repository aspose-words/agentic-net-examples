using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathEnumerationExample
{
    public static void Main()
    {
        // Path for the temporary sample document.
        string samplePath = "SampleEquations.docx";

        // -------------------------------------------------------------
        // 1. Create a sample DOCX document containing a few OfficeMath equations.
        // -------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple fraction equation.
        InsertEquation(builder, @"\f(1,2)");

        // Insert a radical equation.
        InsertEquation(builder, @"\r(3,x)");

        // Insert an integral with a summation.
        InsertEquation(builder, @"\i \su(n=1,5,n)");

        // Save the sample document to disk.
        doc.Save(samplePath);

        // -------------------------------------------------------------
        // 2. Load the document from disk.
        // -------------------------------------------------------------
        Document loadedDoc = new Document(samplePath);

        // -------------------------------------------------------------
        // 3. Enumerate each OfficeMath node (equation) in the document.
        // -------------------------------------------------------------
        NodeCollection officeMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        Console.WriteLine($"Total OfficeMath nodes found: {officeMathNodes.Count}");

        for (int i = 0; i < officeMathNodes.Count; i++)
        {
            OfficeMath om = (OfficeMath)officeMathNodes[i];
            Console.WriteLine($"Equation {i + 1}:");
            Console.WriteLine($"  MathObjectType : {om.MathObjectType}");
            Console.WriteLine($"  DisplayType    : {om.DisplayType}");
            Console.WriteLine($"  Justification  : {om.Justification}");
            Console.WriteLine($"  Text (for reference) : {om.GetText().Trim()}");
        }

        // -------------------------------------------------------------
        // 4. Optional clean‑up: delete the temporary file.
        // -------------------------------------------------------------
        // if (File.Exists(samplePath))
        // {
        //     File.Delete(samplePath);
        // }
    }

    /// <summary>
    /// Inserts an EQ field with the specified arguments, converts it to a real OfficeMath node,
    /// inserts the OfficeMath before the field, and removes the original field.
    /// </summary>
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments after the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        // If conversion succeeded, replace the field with the OfficeMath node.
        if (officeMath != null)
        {
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }

        // Add a new paragraph after the equation for readability.
        builder.InsertParagraph();
    }
}
