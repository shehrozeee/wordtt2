using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace WordTableCellLock
{
    class Program
    {
        static void Main(string[] args)
        {
            //TestLockingScenario();
            string jsonFilePath = Path.Combine(Environment.CurrentDirectory, "validation_lcms_table 2.json");
            string documentPath = Path.Combine(Environment.CurrentDirectory, "UpdatedDocumentForHeader.docx");
            UpdatePropertiesInWordDocument(documentPath, jsonFilePath);
        }

        public class ProjectProperty
        {
            public string Property { get; set; }
            [JsonProperty("Text-Name")]
            public string TextName { get; set; }
        }   

        public static List<ProjectProperty> LoadProjectProperties(string jsonFilePath)
        {
            if (!File.Exists(jsonFilePath))
            {
                throw new FileNotFoundException($"JSON file not found: {jsonFilePath}");
            }

            string jsonContent = File.ReadAllText(jsonFilePath);
            return Newtonsoft.Json.JsonConvert.DeserializeObject<List<ProjectProperty>>(jsonContent);
        }

        public static void UpdatePropertiesInWordDocument(string DocumentPath, string jsonFilePath)
        {
            List<ProjectProperty> properties = LoadProjectProperties(jsonFilePath);
            if (!File.Exists(DocumentPath))
            {
                throw new FileNotFoundException($"Word document not found: {DocumentPath}");
            }

            Word.Application wordApp = null;
            Word.Document doc = null;
            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;
                doc = wordApp.Documents.Open(DocumentPath, ReadOnly: false, Visible: false);

                foreach (var property in properties)
                {
                    ReplaceTextWithDocumentProperty(doc, property.Property, property.TextName, wordApp);
                }
                foreach (Word.Section section in doc.Sections)
                {
                    // Update header fields
                    foreach (Word.HeaderFooter header in section.Headers)
                    {
                        foreach (Word.Field field in header.Range.Fields)
                        {
                            field.Update();
                        }
                    }
                        
                    // Update footer fields
                    foreach (Word.HeaderFooter footer in section.Footers)
                    {
                        foreach (Word.Field field in footer.Range.Fields)
                        {
                            field.Update();
                        }
                    }
                }
                doc.Fields.Update();
                //Save as a new document to avoid overwriting the original
                string outputPath = Path.Combine(Environment.CurrentDirectory, "UpdatedDocument.docx");
                SaveDocument(doc, outputPath, Missing.Value);

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(SaveChanges: true);
                    Marshal.ReleaseComObject(doc);
                }
                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }
            }
        }

        static void TestLockingScenario()
        {
            string inputPath = Path.Combine(Environment.CurrentDirectory, "test.docx");
            string outputPath = Path.Combine(Environment.CurrentDirectory, "test.protected.docx");
            string password = "mypassword";
            object missing = Missing.Value;
            object anEditorID = Word.WdEditorType.wdEditorEveryone;

            Word.Application wordApp = null;
            Word.Document doc = null;
            try
            {
                wordApp = CreateWordApp();
                doc = OpenDocument(wordApp, inputPath);

                // Example: Replace text with document properties
                ReplaceTextWithDocumentProperty(doc, "CompanyName", "[COMPANY]" , wordApp);
                ReplaceTextWithDocumentProperty(doc, "ProjectTitle", "[PROJECT]",wordApp);

                LockExceptBookmarks(doc, new[] { "table1", "table2" }, password, anEditorID, missing, protect: false);
                UnlockTextRange(doc, "[sgfs]", "[sgfe]", anEditorID);
                ProtectDocument(doc, password, missing);
                SaveDocument(doc, outputPath, missing);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                CloseWord(doc, wordApp, missing);
            }
        }

        static Word.Application CreateWordApp()
        {
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            return wordApp;
        }

        static Word.Document OpenDocument(Word.Application wordApp, string path)
        {
            return wordApp.Documents.Open(path, ReadOnly: false, Visible: false);
        }

        static void SaveDocument(Word.Document doc, string path, object missing)
        {
            object filePathObj = path;
            doc.SaveAs2(ref filePathObj, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing);
            Console.WriteLine($"Document saved to {path}");
        }

        static void CloseWord(Word.Document doc, Word.Application wordApp, object missing)
        {
            if (doc != null)
            {
                object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                doc.Close(ref saveChanges, ref missing, ref missing);
                Marshal.ReleaseComObject(doc);
            }
            if (wordApp != null)
            {
                wordApp.Quit(ref missing, ref missing, ref missing);
                Marshal.ReleaseComObject(wordApp);
            }
        }

        static void LockExceptBookmarks(Word.Document doc, string[] lockedBookmarkNames, string password, object anEditorID, object missing, bool protect = true)
        {
            // Unprotect if already protected
            if (doc.ProtectionType != Word.WdProtectionType.wdNoProtection)
            {
                doc.Unprotect(password);
            }
            // 1. Remove all editors from the document (lock everything)
            var docEditors = doc.Range().Editors;
            while (docEditors.Count > 0)
            {
                var editor = docEditors.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, docEditors, new object[] { 1 });
                var deleteMethod = editor.GetType().GetMethod("Delete");
                deleteMethod.Invoke(editor, null);
            }
            // 2. Get all bookmark names and ranges
            var allBookmarkNames = new List<string>();
            var allBookmarkRanges = new List<Word.Range>();
            foreach (Word.Bookmark bm in doc.Bookmarks)
            {
                allBookmarkNames.Add(bm.Name);
                allBookmarkRanges.Add(bm.Range);
            }
            // 3. Unlock all bookmarks except those that should remain locked
            for (int i = 0; i < allBookmarkNames.Count; i++)
            {
                string name = allBookmarkNames[i];
                Word.Range range = allBookmarkRanges[i];
                if (System.Array.IndexOf(lockedBookmarkNames, name) == -1)
                {
                    range.Editors.Add(ref anEditorID);
                    Console.WriteLine($"Bookmark '{name}' marked as editable.");
                }
                else
                {
                    Console.WriteLine($"Bookmark '{name}' will remain locked.");
                }
            }
            // 4. Optionally, unlock content outside all bookmarks
            // Find all editable spans outside locked bookmarks
            int docStart = doc.Content.Start;
            int docEnd = doc.Content.End;
            var lockedRanges = new List<(int, int)>();
            foreach (var name in lockedBookmarkNames)
            {
                if (doc.Bookmarks.Exists(name))
                {
                    var r = doc.Bookmarks[name].Range;
                    lockedRanges.Add((r.Start, r.End));
                }
            }
            lockedRanges.Sort((a, b) => a.Item1.CompareTo(b.Item1));
            int pos = docStart;
            foreach (var pair in lockedRanges)
            {
                int start = pair.Item1;
                int end = pair.Item2;
                if (pos < start)
                {
                    var rng = doc.Range(pos, start);
                    rng.Editors.Add(ref anEditorID);
                }
                pos = end;
            }
            if (pos < docEnd)
            {
                var rng = doc.Range(pos, docEnd);
                rng.Editors.Add(ref anEditorID);
            }
            // 5. Protect the document for tracked changes (Editors only works for tracked changes or comments)
            if (protect)
            {
                ProtectDocument(doc, password, missing);
            }
        }

        static void ProtectDocument(Word.Document doc, string password, object missing)
        {
            object noResetTrue = true;
            Console.WriteLine("Protecting document for tracked changes...");
            doc.Protect(Word.WdProtectionType.wdAllowOnlyReading, ref noResetTrue, password, ref missing, ref missing);
            Console.WriteLine("Document protected. All except specified bookmarks are editable.");
        }

        static void UnlockTextRange(Word.Document doc, string startTag, string endTag, object anEditorID)
        {
            Word.Range contentRange = doc.Content.Duplicate;
            string docText = contentRange.Text;
            int searchStart = 0;
            int unlockCount = 0;
            while (true)
            {
                int startIdx = docText.IndexOf(startTag, searchStart);
                if (startIdx == -1) break;
                int endIdx = docText.IndexOf(endTag, startIdx + startTag.Length);
                if (endIdx == -1) break;
                int rangeStart = contentRange.Start + startIdx + startTag.Length;
                int rangeEnd = contentRange.Start + endIdx;
                if (rangeEnd > rangeStart)
                {
                    Word.Range unlockRange = doc.Range(rangeStart, rangeEnd);
                    unlockRange.Editors.Add(ref anEditorID);
                    unlockCount++;
                    Console.WriteLine($"Content between '{startTag}' and '{endTag}' (instance {unlockCount}) was marked as editable, even if inside a locked region.");
                }
                searchStart = endIdx + endTag.Length;
            }
            if (unlockCount == 0)
            {
                Console.WriteLine($"No content found between tags '{startTag}' and '{endTag}'.");
            }
        }

        static void UnlockBookmarks(Word.Document doc, string[] bookmarkNames, object anEditorID)
        {
            foreach (var name in bookmarkNames)
            {
                if (doc.Bookmarks.Exists(name))
                {
                    var range = doc.Bookmarks[name].Range;
                    range.Editors.Add(ref anEditorID);
                    Console.WriteLine($"Bookmark '{name}' was marked as editable, even if inside a locked region.");
                }
                else
                {
                    Console.WriteLine($"Bookmark '{name}' not found.");
                }
            }
        }

        /// <summary>
        /// Searches for all instances of specified text and replaces them with a document property field.
        /// If the document property doesn't exist, it will be created with the property name as its value.
        /// </summary>
        /// <param name="doc">The Word document to modify</param>
        /// <param name="propertyName">The name of the document property to create/use</param>
        /// <param name="textToMatch">The text to search for and replace</param>
        static void ReplaceTextWithDocumentProperty(Word.Document doc, string propertyName, string textToMatch,Word.Application wordApp)
        {
            try
            {
                object missing = Missing.Value;
                
                // Get custom document properties
                dynamic customProperties = doc.CustomDocumentProperties;
                // Check if property already exists (case-insensitive)
                bool propertyExists = false;
                string existingPropertyName = propertyName;
                
                // Iterate through existing properties
                for (int i = 1; i <= customProperties.Count; i++)
                {
                    var prop = customProperties.Item(i);
                    if (prop.Name.ToString().Equals(propertyName, StringComparison.OrdinalIgnoreCase))
                    {
                        propertyExists = true;
                        existingPropertyName = prop.Name.ToString();
                        break;
                    }
                }
                
                // Add property if it doesn't exist
                if (!propertyExists)
                {
                    customProperties.Add(propertyName, false, 4, propertyName); // 4 = msoPropertyTypeString
                    Console.WriteLine($"Document property '{propertyName}' created with value '{propertyName}'");
                }
                else
                {
                    Console.WriteLine($"Document property '{existingPropertyName}' already exists");
                }
                int replacementCount = 0;

                replacementCount = ReplaceInStoryRange(doc.Content, textToMatch, existingPropertyName, wordApp);
                int numberOfSections = doc.Sections.Count;
                foreach (Word.Section section in doc.Sections)
                {
                    Console.WriteLine($"Processing section {section.Index} of {numberOfSections}");
                    int numberOfHeaders = section.Headers.Count;
                    // Update header fields
                    foreach (Word.HeaderFooter header in section.Headers)
                    {
                        Console.WriteLine($"Processing header {header.Index} of {numberOfHeaders} in section {section.Index}");
                        ReplaceInStoryRange(header.Range, textToMatch, existingPropertyName, wordApp);
                    }

                    int numberOfFooters = section.Footers.Count;
                    // Update footer fields
                    foreach (Word.HeaderFooter footer in section.Footers)
                    {
                        Console.WriteLine($"Processing footer {footer.Index} of {numberOfFooters} in section {section.Index}");
                        ReplaceInStoryRange(footer.Range, textToMatch, existingPropertyName, wordApp);
                    }
                }

                if (replacementCount > 0)
                {
                    Console.WriteLine($"Replaced {replacementCount} instances of '{textToMatch}' with document property field '{existingPropertyName}'");
                }
                else
                {
                    Console.WriteLine($"No instances of '{textToMatch}' found in the document");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in ReplaceTextWithDocumentProperty: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static int ReplaceInStoryRange(Range storyRange, string searchText, string propertyName, Application wordApp)
        {
            int replacementCount = 0;
            if (storyRange == null) return replacementCount;

            Find findObject = storyRange.Find;
            findObject.ClearFormatting();
            findObject.Text = searchText;
            findObject.Forward = true;
            findObject.Wrap = WdFindWrap.wdFindStop;
            Console.WriteLine($"Searching for '{searchText}' in story range...");
            while (findObject.Execute())
            {
                Console.WriteLine($"Found '{searchText}' at position {storyRange.Start}");
                storyRange.Select();
                wordApp.Selection.Delete();

                Field field = storyRange.Document.Fields.Add(wordApp.Selection.Range, WdFieldType.wdFieldDocProperty, propertyName);

                // Reset the story range for next search
                storyRange.Start = wordApp.Selection.Range.End;
                findObject = storyRange.Find;
                findObject.Text = searchText;
                findObject.Forward = true;
                findObject.Wrap = WdFindWrap.wdFindStop;
                replacementCount++;
            }
            return replacementCount;
        }
    }
}