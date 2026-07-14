using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some introductory text.
        builder.Writeln("Below are several equations that will be displayed inline.");

        // Helper to insert an EQ field, convert it to OfficeMath, and clean up the field.
        void InsertEquation(string eqArguments)
        {
            // Insert an EQ field.
            Field field = builder.InsertField(FieldType.FieldEquation, true);
            FieldEQ fieldEq = field as FieldEQ;
            if (fieldEq == null)
                throw new InvalidOperationException("Failed to create FieldEQ.");

            // Write the EQ arguments (e.g., "\f(1,2)").
            builder.MoveTo(fieldEq.Separator);
            builder.Write(eqArguments);

            // Return the builder to the paragraph after the field.
            builder.MoveTo(fieldEq.Start.ParentNode);
            builder.InsertParagraph();

            // Convert the field to a real OfficeMath object.
            OfficeMath officeMath = fieldEq.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                fieldEq.Start.ParentNode.InsertBefore(officeMath, fieldEq.Start);
                // Remove the original field.
                fieldEq.Remove();
            }
        }

        // Insert a few sample equations using safe EQ switches.
        InsertEquation(@"\f(1,2)");          // Fraction 1/2
        InsertEquation(@"\r(3,x)");          // Cube root of x
        InsertEquation(@"\i \su(n=1,5,n)"); // Integral with summation

        // Iterate over all OfficeMath nodes and set their display type to Inline
        // (only for top‑level equations, i.e., MathObjectType.OMathPara).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                om.DisplayType = OfficeMathDisplayType.Inline;
                // Justification cannot be set when DisplayType is Inline, so we leave it unchanged.
            }
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "InlineEquations.docx");
        doc.Save(outputPath);
    }
}
