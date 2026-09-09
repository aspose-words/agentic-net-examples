using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class OfficeMathReportGenerator
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several sample equations using the deterministic EQ‑field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");               // Fraction 1/2
        InsertEquation(builder, @"\r(3,x)");               // Cube root of x
        InsertEquation(builder, @"\i \su(n=1,5,n)");       // Integral with summation
        InsertEquation(builder, @"\s \up8(Superscript)"); // Superscript example

        // Save the document containing the equations.
        string docPath = "OfficeMathSample.docx";
        doc.Save(docPath, SaveFormat.Docx);

        // Prepare the report file.
        string reportPath = "OfficeMathReport.txt";
        using (StreamWriter writer = new StreamWriter(reportPath))
        {
            // Retrieve all OfficeMath nodes in the document (including nested ones).
            NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

            writer.WriteLine($"Total OfficeMath nodes found: {mathNodes.Count}");
            writer.WriteLine();

            // Enumerate each OfficeMath node and write its type and position.
            for (int i = 0; i < mathNodes.Count; i++)
            {
                OfficeMath officeMath = (OfficeMath)mathNodes[i];
                // Position information: index in the collection and the paragraph index.
                int paragraphIndex = officeMath.ParentParagraph?.ParentNode?.GetChildNodes(NodeType.Paragraph, true).IndexOf(officeMath.ParentParagraph) ?? -1;

                writer.WriteLine($"Equation #{i + 1}");
                writer.WriteLine($"  MathObjectType : {officeMath.MathObjectType}");
                writer.WriteLine($"  Collection Index: {i}");
                writer.WriteLine($"  Paragraph Index : {paragraphIndex}");
                writer.WriteLine();
            }
        }

        // Validate that the report was created.
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Report file was not created.");

        Console.WriteLine($"Document saved to '{docPath}'.");
        Console.WriteLine($"Report generated at '{reportPath}'.");
    }

    // Inserts an EQ field with the given argument string, converts it to a real OfficeMath node,
    // inserts the OfficeMath before the field, and removes the original field.
    private static void InsertEquation(DocumentBuilder builder, string eqArgument)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the EQ argument.
        builder.MoveTo(field.Separator);
        builder.Write(eqArgument);

        // Return the builder to the field start's parent (the paragraph).
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to OfficeMath.
        OfficeMath officeMath = field.AsOfficeMath();

        // Ensure conversion succeeded before proceeding.
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field from the document.
            field.Remove();
        }

        // Add a new paragraph after the equation for readability.
        builder.InsertParagraph();
    }
}
