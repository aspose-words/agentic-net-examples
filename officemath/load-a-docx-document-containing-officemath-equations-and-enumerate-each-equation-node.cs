using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Define paths for the sample document and the report.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string docPath = Path.Combine(outputDir, "SampleEquations.docx");
        string reportPath = Path.Combine(outputDir, "EquationsReport.txt");

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX file that contains a few OfficeMath equations.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper to insert an EQ field, convert it to OfficeMath, and clean up.
        void InsertEquation(string eqArgs)
        {
            // Insert an EQ field.
            FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
            // Write the EQ arguments (e.g., "\f(1,2)").
            builder.MoveTo(field.Separator);
            builder.Write(eqArgs);
            // Return the cursor to the paragraph that contains the field.
            builder.MoveTo(field.Start.ParentNode);
            // Convert the field to a real OfficeMath object.
            OfficeMath officeMath = field.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                field.Start.ParentNode.InsertBefore(officeMath, field.Start);
                // Remove the original field.
                field.Remove();
            }
            // Start a new paragraph for the next equation.
            builder.InsertParagraph();
        }

        // Insert several simple equations.
        InsertEquation(@"\f(1,2)");          // Fraction 1/2
        InsertEquation(@"\r(3,x)");          // Cube root of x
        InsertEquation(@"\i \su(n=1,5,n)"); // Integral with summation

        // Save the document to disk.
        doc.Save(docPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the document back from disk.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // -----------------------------------------------------------------
        // 3. Enumerate all OfficeMath nodes in the document.
        // -----------------------------------------------------------------
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);

        using (StreamWriter writer = new StreamWriter(reportPath))
        {
            writer.WriteLine($"Total OfficeMath nodes found: {mathNodes.Count}");
            for (int i = 0; i < mathNodes.Count; i++)
            {
                OfficeMath math = (OfficeMath)mathNodes[i];
                // Output basic information about each equation.
                writer.WriteLine($"Equation {i + 1}:");
                writer.WriteLine($"  MathObjectType: {math.MathObjectType}");
                writer.WriteLine($"  DisplayType: {math.DisplayType}");
                writer.WriteLine($"  Text: {math.GetText().Trim()}");
            }
        }

        // -----------------------------------------------------------------
        // 4. Validate that the output files were created.
        // -----------------------------------------------------------------
        if (!File.Exists(docPath))
            throw new FileNotFoundException("The sample document was not created.", docPath);
        if (!File.Exists(reportPath))
            throw new FileNotFoundException("The report file was not created.", reportPath);

        // The program finishes without waiting for user input.
    }
}
