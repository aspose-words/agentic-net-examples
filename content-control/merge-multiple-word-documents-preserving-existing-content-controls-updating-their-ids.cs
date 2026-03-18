using System;
using System.Collections.Generic;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Markup;

namespace AsposeWordsMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Paths of the source documents to be merged.
            string[] sourceFiles = new[]
            {
                @"C:\Docs\Part1.docx",
                @"C:\Docs\Part2.docx",
                @"C:\Docs\Part3.docx"
            };

            // Path of the resulting merged document.
            string outputFile = @"C:\Docs\MergedResult.docx";

            // Create a blank destination document.
            Document mergedDoc = new Document();
            mergedDoc.EnsureMinimum(); // Ensure at least one section.

            // Append each source document while keeping its original formatting.
            foreach (string file in sourceFiles)
            {
                Document src = new Document(file);
                mergedDoc.AppendDocument(src, ImportFormatMode.KeepSourceFormatting);
            }

            // After merging, reassign unique IDs to all content controls.
            ReassignContentControlIds(mergedDoc);

            // Save the merged document.
            mergedDoc.Save(outputFile);
        }

        /// <summary>
        /// Assigns new sequential IDs to every StructuredDocumentTag (content control) in the document.
        /// Uses reflection because the Id property is read‑only in older Aspose.Words versions.
        /// </summary>
        private static void ReassignContentControlIds(Document doc)
        {
            NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
            int nextId = 1;
            // Cache the PropertyInfo for performance.
            PropertyInfo idProperty = typeof(StructuredDocumentTag).GetProperty("Id", BindingFlags.Instance | BindingFlags.Public);

            foreach (StructuredDocumentTag sdt in sdtNodes)
            {
                // If the Id property has a setter (newer versions) we can assign directly.
                // Otherwise, fall back to reflection to set the backing field.
                if (idProperty != null && idProperty.CanWrite)
                {
                    idProperty.SetValue(sdt, nextId);
                }
                else
                {
                    // For read‑only Id, set the underlying private field "_id" via reflection.
                    FieldInfo idField = typeof(StructuredDocumentTag).GetField("_id", BindingFlags.Instance | BindingFlags.NonPublic);
                    idField?.SetValue(sdt, nextId);
                }
                nextId++;
            }
        }
    }
}
