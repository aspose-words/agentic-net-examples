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
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Local helper that inserts an equation using the deterministic EQ-field bootstrap workflow.
        void InsertEquation(string eqSwitch)
        {
            // Insert an empty Equation field.
            Field field = builder.InsertField(FieldType.FieldEquation, true);

            // Move the builder to the field separator where the EQ argument should be written.
            builder.MoveTo(field.Separator);
            builder.Write(eqSwitch);

            // Convert the field to a real OfficeMath node.
            FieldEQ fieldEq = field as FieldEQ;
            OfficeMath officeMath = fieldEq?.AsOfficeMath();

            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                Node fieldStart = field.Start;
                fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);

                // Remove the original field so only the OfficeMath remains.
                field.Remove();
            }
        }

        // Add some paragraphs and equations.
        builder.Writeln("Paragraph before equations.");
        InsertEquation(@"\f(1,2)"); // Simple fraction.
        builder.Writeln(); // New line after the equation.
        InsertEquation(@"\r(3,x)"); // Simple root.
        builder.Writeln(); // New line after the equation.
        InsertEquation(@"\s(5)"); // Simple sum.
        builder.Writeln("Paragraph after equations.");

        // Save the document.
        string docPath = "Sample.docx";
        doc.Save(docPath, SaveFormat.Docx);

        // Reload the document to ensure proper loading.
        Document loadedDoc = new Document(docPath);

        // Retrieve all OfficeMath nodes.
        NodeCollection officeMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);

        // Prepare report lines.
        List<string> reportLines = new List<string>();
        reportLines.Add("Index,MathObjectType,ParagraphIndex");

        // Retrieve all paragraphs for index calculation.
        NodeCollection paragraphs = loadedDoc.GetChildNodes(NodeType.Paragraph, true);

        for (int i = 0; i < officeMathNodes.Count; i++)
        {
            OfficeMath om = (OfficeMath)officeMathNodes[i];
            Paragraph para = om.ParentNode as Paragraph;
            int paraIndex = -1;
            if (para != null)
            {
                paraIndex = paragraphs.IndexOf(para);
            }

            string line = $"{i + 1},{om.MathObjectType},{paraIndex}";
            reportLines.Add(line);
        }

        // Write the report to a text file.
        string reportPath = "OfficeMathReport.txt";
        File.WriteAllLines(reportPath, reportLines);

        // Validate that the report file was created.
        if (!File.Exists(reportPath))
        {
            throw new Exception("Report file was not created.");
        }
    }
}
