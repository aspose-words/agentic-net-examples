using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathEnumerationExample
{
    public static void Main()
    {
        // Prepare output folder.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);
        string docPath = Path.Combine(dataDir, "Equations.docx");
        string reportPath = Path.Combine(dataDir, "EquationsReport.txt");

        // 1. Create a sample DOCX containing a few equations.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three simple equations using the EQ field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");          // Fraction 1/2
        InsertEquation(builder, @"\r(3,x)");          // Cube root of x
        InsertEquation(builder, @"\i \su(n=1,5,n)"); // Integral with summation

        // Save the sample document.
        doc.Save(docPath);

        // 2. Load the document that now contains OfficeMath nodes.
        Document loadedDoc = new Document(docPath);

        // 3. Enumerate all OfficeMath nodes in the document.
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        StringBuilder report = new StringBuilder();

        for (int i = 0; i < mathNodes.Count; i++)
        {
            OfficeMath om = (OfficeMath)mathNodes[i];
            report.AppendLine($"Equation {i + 1}: MathObjectType={om.MathObjectType}, DisplayType={om.DisplayType}");
        }

        // Write the enumeration report to a text file.
        File.WriteAllText(reportPath, report.ToString());

        // Also output the report to the console.
        Console.WriteLine(report.ToString());

        // Validate that the report file was created.
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Report file was not created.");
    }

    // Helper method that inserts an EQ field, converts it to a real OfficeMath node, and cleans up the field.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start node.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field, leaving only the OfficeMath node.
            field.Remove();
        }

        // Add a new paragraph after the equation for readability.
        builder.InsertParagraph();
    }
}
