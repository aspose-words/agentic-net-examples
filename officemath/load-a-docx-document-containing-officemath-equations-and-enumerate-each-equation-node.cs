using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathEnumerationExample
{
    public static void Main()
    {
        // Define file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string docPath = Path.Combine(outputDir, "SampleEquations.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document with a few OfficeMath equations.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper to insert an EQ field, convert it to OfficeMath and clean up.
        void InsertEquation(string eqArguments)
        {
            // Insert an EQ field.
            FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
            // Write the EQ arguments into the field separator.
            builder.MoveTo(field.Separator);
            builder.Write(eqArguments);
            // Return the builder to the paragraph that contains the field.
            builder.MoveTo(field.Start.ParentNode);

            // Convert the EQ field to a real OfficeMath object.
            OfficeMath officeMath = field.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                field.Start.ParentNode.InsertBefore(officeMath, field.Start);
                // Remove the original field.
                field.Remove();
            }

            // Add a paragraph break after each equation for readability.
            builder.Writeln();
        }

        // Insert a few simple equations using safe EQ switches.
        InsertEquation(@"\f(1,2)");          // Fraction 1/2
        InsertEquation(@"\r(3,x)");          // Cube root of x
        InsertEquation(@"\i \su(n=1,5,n)"); // Integral with summation

        // Save the sample document.
        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new InvalidOperationException("Failed to create the sample document.");

        // ---------------------------------------------------------------
        // 2. Load the document and enumerate all OfficeMath nodes.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);

        Console.WriteLine($"Total OfficeMath nodes found: {mathNodes.Count}");
        for (int i = 0; i < mathNodes.Count; i++)
        {
            OfficeMath math = (OfficeMath)mathNodes[i];
            // Output basic information about each equation.
            Console.WriteLine($"Equation {i + 1}:");
            Console.WriteLine($"  MathObjectType : {math.MathObjectType}");
            Console.WriteLine($"  DisplayType    : {math.DisplayType}");
            Console.WriteLine($"  Text           : {math.GetText().Trim()}");
        }

        // Optional: write a simple text report.
        string reportPath = Path.Combine(outputDir, "EquationReport.txt");
        using (StreamWriter writer = new StreamWriter(reportPath))
        {
            writer.WriteLine($"Document: {Path.GetFileName(docPath)}");
            writer.WriteLine($"OfficeMath nodes count: {mathNodes.Count}");
            for (int i = 0; i < mathNodes.Count; i++)
            {
                OfficeMath math = (OfficeMath)mathNodes[i];
                writer.WriteLine($"Equation {i + 1}: Type={math.MathObjectType}, Display={math.DisplayType}, Text=\"{math.GetText().Trim()}\"");
            }
        }

        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Failed to create the report file.");
    }
}
