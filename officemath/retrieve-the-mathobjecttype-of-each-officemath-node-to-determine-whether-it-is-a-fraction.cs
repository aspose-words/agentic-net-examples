using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Fields;

public class OfficeMathTypeDemo
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Insert a fraction equation ----------
        // Insert an EQ field.
        FieldEQ fractionField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Move to the field separator and write the EQ switch for a fraction.
        builder.MoveTo(fractionField.Separator);
        builder.Write(@"\f(1,2)");
        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(fractionField.Start.ParentNode);

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath fractionMath = fractionField.AsOfficeMath();
        if (fractionMath != null)
        {
            // Insert the OfficeMath before the field start and remove the original field.
            fractionField.Start.ParentNode.InsertBefore(fractionMath, fractionField.Start);
            fractionField.Remove();
        }

        // Insert a paragraph break between equations.
        builder.Writeln();

        // ---------- Insert a radical equation ----------
        FieldEQ radicalField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        builder.MoveTo(radicalField.Separator);
        builder.Write(@"\r(3,x)");
        builder.MoveTo(radicalField.Start.ParentNode);

        OfficeMath radicalMath = radicalField.AsOfficeMath();
        if (radicalMath != null)
        {
            radicalField.Start.ParentNode.InsertBefore(radicalMath, radicalField.Start);
            radicalField.Remove();
        }

        // Save the document (optional, demonstrates lifecycle).
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OfficeMathTypes.docx");
        doc.Save(outputPath);

        // ---------- Enumerate OfficeMath nodes and report their types ----------
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        for (int i = 0; i < mathNodes.Count; i++)
        {
            OfficeMath om = (OfficeMath)mathNodes[i];
            MathObjectType type = om.MathObjectType;

            string typeDescription = type switch
            {
                MathObjectType.Fraction => "Fraction",
                MathObjectType.Radical => "Radical",
                _ => type.ToString()
            };

            Console.WriteLine($"Equation {i}: {typeDescription}");
        }
    }
}
