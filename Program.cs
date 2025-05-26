using System;
using System.Collections.Generic;
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
            TestLockingScenario();
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
                LockExceptBookmarks(doc, new[] { "table1", "table2" }, password, anEditorID, missing, protect: false);
                UnlockTextRange(doc, "Cat", anEditorID);
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

        static void UnlockTextRange(Word.Document doc, string text, object anEditorID)
        {
            Word.Range searchRange = doc.Content.Duplicate;
            Word.Find find = searchRange.Find;
            find.Text = text;
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            if (find.Execute())
            {
                searchRange.Editors.Add(ref anEditorID);
                Console.WriteLine($"The text '{text}' was found and marked as editable, even if inside a locked region.");
            }
            else
            {
                Console.WriteLine($"The text '{text}' was not found in the document.");
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
    }
}