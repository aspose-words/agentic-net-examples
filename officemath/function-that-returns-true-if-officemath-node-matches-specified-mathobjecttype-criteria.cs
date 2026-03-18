using System;
using Aspose.Words;
using Aspose.Words.Math;

namespace OfficeMathUtilities
{
    public static class OfficeMathHelper
    {
        /// <summary>
        /// Determines whether the given <see cref="OfficeMath"/> node has the specified <see cref="MathObjectType"/>.
        /// Returns false if the node is null.
        /// </summary>
        /// <param name="officeMath">The OfficeMath node to examine.</param>
        /// <param name="expectedType">The MathObjectType to compare against.</param>
        /// <returns>True if the node's MathObjectType matches the expected type; otherwise false.</returns>
        public static bool IsOfType(OfficeMath officeMath, MathObjectType expectedType)
        {
            // Guard against null references.
            if (officeMath == null)
                return false;

            // Compare the MathObjectType property with the supplied criteria.
            return officeMath.MathObjectType == expectedType;
        }
    }

    // Simple console entry point required for a .NET executable project.
    internal class Program
    {
        static void Main(string[] args)
        {
            // Demonstration – this block can be removed or expanded as needed.
            try
            {
                // Load a document that contains OfficeMath objects.
                Document doc = new Document("OfficeMath.docx");
                // Retrieve the first OfficeMath node (if any).
                OfficeMath om = doc.GetChild(NodeType.OfficeMath, 0, true) as OfficeMath;
                // Check its type against a specific MathObjectType.
                bool isPara = OfficeMathHelper.IsOfType(om, MathObjectType.OMathPara);
                Console.WriteLine($"Is OMathPara: {isPara}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during example execution: {ex.Message}");
            }
        }
    }
}
