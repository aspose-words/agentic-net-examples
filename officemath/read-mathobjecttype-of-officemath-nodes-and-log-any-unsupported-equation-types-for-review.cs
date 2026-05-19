using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Fields;

public class OfficeMathTypeReporter
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several equations using the deterministic EQ‑field bootstrap workflow.
        // 1. Simple fraction: \f(1,2)
        InsertAndConvertEquation(builder, @"\f(1,2)");

        // 2. Radical (cube root): \r(3,x)
        InsertAndConvertEquation(builder, @"\r(3,x)");

        // 3. Array (2×2 matrix) with nested fractions.
        InsertAndConvertEquation(builder, @"\a \co2 \vs1 \hs1( \f(1,2), \f(3,4), \f(5,6), \f(7,8) )");

        // 4. Integral with limits: \i \su(n=1,5,n)
        InsertAndConvertEquation(builder, @"\i \su(n=1,5,n)");

        // Save the document (optional, for visual inspection).
        string docPath = Path.Combine(Environment.CurrentDirectory, "OfficeMathTypes.docx");
        doc.Save(docPath);

        // Enumerate all OfficeMath nodes in the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        List<string> unsupportedReports = new List<string>();

        for (int i = 0; i < mathNodes.Count; i++)
        {
            OfficeMath officeMath = (OfficeMath)mathNodes[i];
            MathObjectType type = officeMath.MathObjectType;

            // Consider OMathPara (top‑level equation) as supported; everything else is logged.
            if (type != MathObjectType.OMathPara)
            {
                string report = $"Unsupported MathObjectType: {type} (Node index {i})";
                Console.WriteLine(report);
                unsupportedReports.Add(report);
            }
        }

        // Write a simple text report file if any unsupported types were found.
        if (unsupportedReports.Count > 0)
        {
            string reportPath = Path.Combine(Environment.CurrentDirectory, "UnsupportedMathTypes.txt");
            File.WriteAllLines(reportPath, unsupportedReports);
            Console.WriteLine($"Report written to: {reportPath}");
        }
        else
        {
            Console.WriteLine("All OfficeMath nodes are of supported type OMathPara.");
        }
    }

    // Inserts an EQ field with the given arguments, converts it to OfficeMath, and cleans up the field.
    private static void InsertAndConvertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert the EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Write the arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);
        // Return to the paragraph that contains the field and start a new paragraph for the next equation.
        builder.MoveTo(field.Start.ParentNode);
        builder.InsertParagraph();

        // Convert the field to a real OfficeMath node.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field.
            field.Remove();
        }
    }
}
