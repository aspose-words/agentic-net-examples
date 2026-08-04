using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathInspector
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string docPath = Path.Combine(artifactsDir, "OfficeMathSample.docx");

        // Create a new document and insert a few simple equations using the EQ field bootstrap workflow.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        InsertAndConvert(builder, @"\f(1,2)");               // Fraction 1/2
        InsertAndConvert(builder, @"\r(3,x)");               // Cube root of x
        InsertAndConvert(builder, @"\i \su(n=1,5,n)");       // Integral with summation

        // Save the document.
        doc.Save(docPath);

        // Reload the document to simulate a typical load‑process scenario.
        Document loadedDoc = new Document(docPath);

        // Collect all OfficeMath nodes.
        NodeCollection officeMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        List<string> unsupported = new List<string>();

        for (int i = 0; i < officeMathNodes.Count; i++)
        {
            OfficeMath om = (OfficeMath)officeMathNodes[i];

            // Consider only OMathPara (top‑level equations) as supported.
            if (om.MathObjectType != MathObjectType.OMathPara)
            {
                string message = $"Unsupported MathObjectType: {om.MathObjectType} (Node index {i})";
                Console.WriteLine(message);
                unsupported.Add(message);
            }
        }

        // Write a simple report file with the unsupported types.
        string reportPath = Path.Combine(artifactsDir, "UnsupportedMathTypes.txt");
        File.WriteAllLines(reportPath, unsupported);
        Console.WriteLine($"Report written to {reportPath}");
    }

    // Inserts an EQ field with the given arguments, converts it to a real OfficeMath node,
    // and removes the original field.
    private static void InsertAndConvert(DocumentBuilder builder, string args)
    {
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        builder.MoveTo(field.Separator);
        builder.Write(args);
        builder.MoveTo(field.Start.ParentNode);

        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start and then delete the field.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }

        // Add a paragraph break after each equation for readability.
        builder.InsertParagraph();
    }
}
