using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several equations using the deterministic EQ-field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");   // Fraction 1/2
        InsertEquation(builder, @"\r(3,x)"); // Cube root of x
        InsertEquation(builder, @"\i \su(n=1,5,n)"); // Integral with summation
        InsertEquation(builder, @"\s \up8(Sup) \s \do8(Sub)"); // Superscript and subscript

        // Perform bulk updates on all top‑level OfficeMath nodes.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in mathNodes)
        {
            // Change display format and justification.
            om.DisplayType = OfficeMathDisplayType.Display;
            om.Justification = OfficeMathJustification.Left;
        }

        // Validate that each OfficeMath node has the expected MathObjectType (OMathPara for top‑level equations).
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType != MathObjectType.OMathPara)
                throw new Exception($"Unexpected MathObjectType: {om.MathObjectType}");
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved.");

        // Indicate successful completion.
        Console.WriteLine("Document saved and all OfficeMath objects validated successfully.");
    }

    // Helper that inserts an EQ field, converts it to a real OfficeMath node, and returns the node.
    private static OfficeMath InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the equation arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return to the paragraph containing the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start and remove the original field.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }

        // Add a paragraph break after the equation for readability.
        builder.Writeln();

        return officeMath;
    }
}
