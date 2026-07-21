using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few sample equations using the deterministic EQ-field workflow.
        InsertOfficeMath(builder, @"\f(1,2)");               // Fraction 1/2
        InsertOfficeMath(builder, @"\r(3,x)");               // Cube root of x
        InsertOfficeMath(builder, @"\i \su(n=1,5,n)");       // Integral with summation

        // Save the sample document.
        const string docPath = "Sample.docx";
        doc.Save(docPath);

        // Reload the document (optional, demonstrates load workflow).
        Document loadedDoc = new Document(docPath);

        // Collect all top‑level OfficeMath equations.
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        List<string> equations = new List<string>();
        foreach (OfficeMath om in mathNodes)
        {
            // Consider only top‑level math paragraphs (OMathPara).
            if (om.MathObjectType == MathObjectType.OMathPara)
                equations.Add(om.GetText().Trim());
        }

        // Write the extracted equations to a text file.
        const string txtPath = "Equations.txt";
        File.WriteAllLines(txtPath, equations);

        // Simple validation to ensure the output file was created.
        if (!File.Exists(txtPath))
            throw new InvalidOperationException($"Failed to create the report file '{txtPath}'.");
    }

    // Inserts an EQ field, converts it to a real OfficeMath node, and removes the field.
    private static void InsertOfficeMath(DocumentBuilder builder, string eqSwitch)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ switch arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqSwitch);

        // Return the builder to the paragraph containing the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field.
            field.Remove();
        }

        // Add a new paragraph after the equation for subsequent content.
        builder.InsertParagraph();
    }
}
