using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;

namespace OfficeMathInspection
{
    // Visitor that records unsupported MathObjectType values.
    class UnsupportedMathVisitor : DocumentVisitor
    {
        // Define which MathObjectTypes are considered supported.
        // NOTE: The Superscript enum value is not present in older Aspose.Words versions,
        // so it has been removed from the supported list to keep the code compatible.
        private static readonly HashSet<MathObjectType> SupportedTypes = new HashSet<MathObjectType>
        {
            MathObjectType.OMath,          // Inline math text
            MathObjectType.OMathPara,      // Display math paragraph
            MathObjectType.Fraction,       // Fraction
            MathObjectType.Accent,         // Accent
            MathObjectType.Bar,            // Bar
            MathObjectType.Matrix,         // Matrix
            MathObjectType.Radical,        // Radical
            MathObjectType.Function,       // Function
            MathObjectType.Limit,          // Limit
            MathObjectType.Subscript,      // Subscript
            // MathObjectType.Superscript, // <-- removed – not defined in this version of Aspose.Words
        };

        // List to store descriptions of unsupported nodes.
        private readonly List<string> _unsupported = new List<string>();

        // Called when an OfficeMath node is encountered.
        public override VisitorAction VisitOfficeMathStart(OfficeMath officeMath)
        {
            // If the node's type is not in the supported set, record it.
            if (!SupportedTypes.Contains(officeMath.MathObjectType))
            {
                int index = officeMath.ParentNode?.IndexOf(officeMath) ?? -1;
                _unsupported.Add($"Unsupported MathObjectType: {officeMath.MathObjectType} (Parent index: {index})");
            }

            return VisitorAction.Continue;
        }

        // Returns the collected log entries.
        public IEnumerable<string> GetLog() => _unsupported;
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // Create a sample document.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Sample document with a math object (none added).");

            // -----------------------------------------------------------------
            // Run the visitor to detect unsupported MathObjectTypes.
            // -----------------------------------------------------------------
            var visitor = new UnsupportedMathVisitor();
            doc.Accept(visitor);

            // Output any unsupported equation types.
            foreach (var entry in visitor.GetLog())
                Console.WriteLine(entry);

            // -----------------------------------------------------------------
            // Save the document (ensuring the output directory exists).
            // -----------------------------------------------------------------
            string outputDir = "ArtifactsDir";
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "DocumentWithMath_Inspected.docx");
            doc.Save(outputPath);
        }
    }
}
