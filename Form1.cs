using Microsoft.Office.Interop.Word;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Replace_Docx
{
    public partial class frmMain : Form
    {
        // Array of DOCX Files found in Folder
        string[] DOCXfiles;

        // Store slected path of Folder browser dialog in variable
        string selected_path;

        // Create fileCount to counting number of DOCX files found
        int fileCount = 0;

        public frmMain() => InitializeComponent();

        // Handle Methode Search in all Sub-Directory and Get all DOCX files found,
        // and bring out to the string array
        private int SearchDOCXFiles(string path, out string[] DOCXfiles)
        {
            DOCXfiles = Directory
                        .GetFiles(path, "*.*", SearchOption.AllDirectories)
                        .Where(s => s.ToLower().EndsWith(".doc") || s.ToLower().EndsWith(".docx"))
                        .ToArray();
            return DOCXfiles.Length;
        }

        // Methode Write exceptions into log file
        static void LogException(string logFilePath, string filePath, string ex)
        {
            using (StreamWriter writer = new StreamWriter(logFilePath, true))
            {
                string filename = Path.GetFileNameWithoutExtension(filePath);
                writer.WriteLine(filename + " : " + ex);
            }
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FD = new FolderBrowserDialog();
            if (selected_path != null)
                FD.SelectedPath = selected_path;
            if (FD.ShowDialog() == DialogResult.OK)
            {
                string path = FD.SelectedPath;
                selected_path = path;
                TxtBoxLoad.Text = path;
                fileCount = SearchDOCXFiles(path, out DOCXfiles);
                // Check the Empty Folder
                labelInfo.Text = fileCount == 0 ? "Your Folder is Empty." : fileCount + " DOCX files found.";
                labelErrorMessage.Text = "";
            }
        }

        private void GitLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // Go to Github repository
            string url = "https://github.com/abdessalam-aadel/Replace_Docx";

            // Open the URL in the default web browser
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true // Ensures the URL is opened in the default web browser
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }

        private void frmMain_DragDrop(object sender, DragEventArgs e)
        {
            string path = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
            // Condition >> Drag Folder
            if (Directory.Exists(path))
            {
                TxtBoxLoad.Text = path;
                fileCount = SearchDOCXFiles(path, out DOCXfiles);
                selected_path = path;
                // Check the Empty Folder
                labelInfo.Text = fileCount == 0 ? "Your Folder is Empty." : fileCount + " DOCX files found.";
                labelErrorMessage.Text = "";
            }
        }

        private void frmMain_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
            TxtBoxLoad.Text = "Chose your folder location ...";
            labelInfo.Text = "...";
            labelErrorMessage.Text = "";
            DOCXfiles = null;
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (DOCXfiles == null || string.IsNullOrEmpty(TxtBoxLoad.Text))
            {
                labelErrorMessage.Text = "No source folder was selected, Please select one.";
                return;
            }

            else if (DOCXfiles.Length == 0)
            {
                labelErrorMessage.Text = "No DOCX file was found in the selected folder";
                return;
            }

            labelErrorMessage.Text = "";
            Cursor = Cursors.WaitCursor;
            labelInfo.Text = "Processing ...";
            labelErrorMessage.Text = "";

            // Create a new instance of Microsoft Word through the Interop library
            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;

            // Log file
            string logFilePath = selected_path + @"\exceptions.log";
            // Delete the log file if it exists
            if (File.Exists(logFilePath))
            {
                File.Delete(logFilePath);
            }

            foreach (string file in DOCXfiles)
            {
                try
                {
                    wordApp.Visible = false;
                    object missing = System.Reflection.Missing.Value;
                    // Open the document 
                    doc = wordApp.Documents.Open(file);

                    // Loop through all content controls in the document
                    foreach (ContentControl contentControl in doc.ContentControls)
                    {
                        contentControl.Delete(); // Removes the content control
                    }

                    // Track if replacements occurred
                    // Replace text and count replacements
                    var (replacementsMade, totalOccurrences, textExists) = FindAndReplace(wordApp, txtBoxOld.Text, txtBoxNew.Text);

                    var (replacementsMade2, totalOccurrences2, textExists2) = FindAndReplace(wordApp, txtBoxOld2.Text, txtBoxNew2.Text);

                    var (replacementsMade3, totalOccurrences3, textExists3) = FindAndReplace(wordApp, txtBoxOld3.Text, txtBoxNew3.Text);

                    // Save and close
                    doc.Save();

                    // Log the results with replacement count
                    if (replacementsMade > 0 || replacementsMade2 > 0 || replacementsMade3 > 0 || textExists == true || textExists2 == true || textExists3 == true )
                    {
                        LogException(logFilePath, file, "1er : " + replacementsMade + " Replacement was successful.");
                        LogException(logFilePath, file, "2eme : " + replacementsMade2 + " Replacement was successful.");
                        LogException(logFilePath, file, "3eme : " + replacementsMade3 + " Replacement was successful.");
                    }
                    else
                    {
                        LogException(logFilePath, file, "1er : " + replacementsMade + " No replacements were made.");
                        LogException(logFilePath, file, "2eme : " + replacementsMade2 + " No replacements were made.");
                        LogException(logFilePath, file, "3eme : " + replacementsMade3 + " No replacements were made.");
                    }
                }
                catch (Exception ex)
                {
                    // Write Exception into exceptions.log
                    LogException(logFilePath, file, ex.Message);
                    // Continue to the next iteration
                    continue;
                }

                finally
                {
                    // Always close the document properly
                    if (doc != null)
                    {
                        doc.Close();
                        Marshal.ReleaseComObject(doc);
                        doc = null;
                    }
                }
            }

            // Quit the Word application
            wordApp.Quit();
            Marshal.ReleaseComObject(wordApp);
            wordApp = null;

            // Clear string array
            DOCXfiles = null;
            Cursor = Cursors.Default;
            TxtBoxLoad.Text = "Chose your folder location ...";
            labelInfo.Text = "Done.";
        }
        // Methode Find and Replace
        private static (int replacementCount, int totalCount, bool textExists) FindAndReplace(Word.Application wordApp, string findText, string replaceText)
        {
            int replacementCount = 0;
            int totalCount = 0;  // Counter for all found occurrences

            Word.Document doc = wordApp.ActiveDocument;

            // Get the Find object from Word
            Word.Find findObject = wordApp.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = findText;
            findObject.Replacement.ClearFormatting();

            //Check if the replaceText contains special characters like "^p" and handle it correctly
            if (replaceText.Contains("^p"))
            {
                // Replace "^p" with the actual Word paragraph mark (using Word's built-in special character)
                replaceText = replaceText.Replace("^p", "\r");  // "\r" is the Word paragraph mark (carriage return)
            }

            findObject.Replacement.Text = replaceText;

            findObject.MatchCase = true; // For case-sensitive search
            findObject.MatchWholeWord = true; // For matching whole words only

            //Set to replace all occurrences
            object replaceAll = Word.WdReplace.wdReplaceAll;

            // Check if the text exists by executing the find operation
            bool textExists = findObject.Execute();

            if (findText == "")
                return (replacementCount, totalCount, textExists);


            // If text is found, count occurrences and perform replacements
            if (textExists)
            {
                // Count the occurrences of the findText (without replacing)
                while (findObject.Execute())
                {
                    totalCount++;  // Increment total count without replacing
                }

                // Perform the replacements
                while (true)
                {
                    bool found = findObject.Execute(Replace: ref replaceAll);
                    if (!found)
                    {
                        break;  // Stop if no more occurrences are found
                    }

                    replacementCount++;  // Increment the count each time a replacement is made
                }
            }

            // --- PART 2: Replace in text boxes / shapes with TextFrames ---
            foreach (Word.Shape shape in doc.Shapes)
            {
                if (shape.TextFrame.HasText != 0)
                {
                    string shapeText = shape.TextFrame.TextRange.Text;
                    if (shapeText.Contains(findText))
                    {
                        totalCount++;
                        shapeText = shapeText.Replace(findText, replaceText);
                        shape.TextFrame.TextRange.Text = shapeText;
                        replacementCount++;
                        textExists = true;
                    }
                }
            }

            return (replacementCount, totalCount, textExists);

        }

        private static int CountReplacements(Word.Application wordApp, string findText)
        {
            int count = 0;
            Word.Find findObject = wordApp.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = findText;

            // Use a Range to search through the document
            Word.Range documentRange = wordApp.ActiveDocument.Content;
            while (findObject.Execute())
            {
                count++;
            }
            return count;
        }

        private void btnDetect_Click(object sender, EventArgs e)
        {
            if (DOCXfiles == null || string.IsNullOrEmpty(TxtBoxLoad.Text))
            {
                labelErrorMessage.Text = "No source folder was selected, Please select one.";
                return;
            }

            else if (DOCXfiles.Length == 0)
            {
                labelErrorMessage.Text = "No DOCX file was found in the selected folder";
                return;
            }

            labelErrorMessage.Text = "";
            Cursor = Cursors.WaitCursor;
            labelInfo.Text = "Processing ...";
            labelErrorMessage.Text = "";

            // Create a new instance of Microsoft Word through the Interop library
            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;

            // Log file
            string logFilePathCount = selected_path + @"\Detect.log";

            // Delete the log file if it exists
            if (File.Exists(logFilePathCount))
            {
                File.Delete(logFilePathCount);
            }

            foreach (string file in DOCXfiles)
            {
                try
                {
                    wordApp.Visible = false;
                    object missing = System.Reflection.Missing.Value;
                    // Open the document 
                    doc = wordApp.Documents.Open(file);

                    // Get the total number of pages in the document
                    object what = WdGoToItem.wdGoToPage;
                    object which = WdGoToDirection.wdGoToLast;
                    Range lastPageRange = wordApp.Selection.GoTo(ref what, ref which);

                    int lastPageNumber = lastPageRange.Information[WdInformation.wdActiveEndPageNumber];

                    string lastPageContent = "";
                    bool onLastPage = false;

                    foreach (Paragraph para in doc.Paragraphs)
                    {
                        Range paraRange = para.Range;
                        int paraPage = paraRange.Information[WdInformation.wdActiveEndPageNumber];

                        if (paraPage == lastPageNumber)
                        {
                            lastPageContent += paraRange.Text;
                            onLastPage = true;
                        }
                        else if (onLastPage && paraPage > lastPageNumber)
                        {
                            break; // We're past the last page (shouldn't happen, but safe guard)
                        }
                    }

                    // Normalize text (trim spaces)
                    lastPageContent = lastPageContent.Trim();

                    if (lastPageContent == txtBoxSearchText.Text)
                        LogException(logFilePathCount, file, " : The last page contains ONLY "+ txtBoxSearchText.Text);
                    
                    //else
                    //    LogException(logFilePathCount, file, "The last page does NOT match the expected text. Actual content on the last page: " + $"\"{lastPageContent}\"");
                }
                catch (Exception ex)
                {
                    // Write Exception into exceptions.log
                    LogException(logFilePathCount, file, ex.Message);
                    // Continue to the next iteration
                    continue;
                }

                finally
                {
                    // Always close the document properly
                    if (doc != null)
                    {
                        doc.Close();
                        Marshal.ReleaseComObject(doc);
                        doc = null;
                    }
                }
            }

            // Quit the Word application
            wordApp.Quit();
            Marshal.ReleaseComObject(wordApp);
            wordApp = null;

            // Clear string array
            DOCXfiles = null;
            Cursor = Cursors.Default;
            TxtBoxLoad.Text = "Chose your folder location ...";
            labelInfo.Text = "Done.";
        }

        private void btnShow_Click(object sender, EventArgs e)
        {
            if (DOCXfiles == null || string.IsNullOrEmpty(TxtBoxLoad.Text))
            {
                labelErrorMessage.Text = "No source folder was selected, Please select one.";
                return;
            }

            else if (DOCXfiles.Length == 0)
            {
                labelErrorMessage.Text = "No DOCX file was found in the selected folder";
                return;
            }

            labelErrorMessage.Text = "";
            Cursor = Cursors.WaitCursor;
            labelInfo.Text = "Processing ...";
            labelErrorMessage.Text = "";

            // Create a new instance of Microsoft Word through the Interop library
            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;

            // Log file
            string logFilePathShow = selected_path + @"\Show_Content.log";

            // Delete the log file if it exists
            if (File.Exists(logFilePathShow))
            {
                File.Delete(logFilePathShow);
            }

            foreach (string file in DOCXfiles)
            {
                try
                {
                    wordApp.Visible = false;
                    object missing = System.Reflection.Missing.Value;
                    // Open the document 
                    doc = wordApp.Documents.Open(file);

                    // Get the total number of pages in the document
                    object what = WdGoToItem.wdGoToPage;
                    object which = WdGoToDirection.wdGoToLast;
                    Range lastPageRange = wordApp.Selection.GoTo(ref what, ref which);

                    int lastPageNumber = lastPageRange.Information[WdInformation.wdActiveEndPageNumber];

                    string lastPageContent = "";

                    foreach (Paragraph para in doc.Paragraphs)
                    {
                        Range paraRange = para.Range;
                        lastPageContent += paraRange.Text;
                    }

                    Word.Bookmarks bookmarks = doc.Bookmarks;
                    if (bookmarks.Exists("MyBookmark"))
                    {
                        Word.Range bookmarkRange = bookmarks["MyBookmark"].Range;
                        LogException(logFilePathShow, file, "Bookmark : " + Environment.NewLine + bookmarkRange.Text);
                    }

                    foreach (Word.Shape shape in doc.Shapes)
                    {
                        if (shape.TextFrame.HasText != 0)
                        {
                            LogException(logFilePathShow, file, "TextFrame : " + Environment.NewLine + shape.TextFrame.TextRange.Text);
                        }
                    }

                    LogException(logFilePathShow, file, Environment.NewLine + $"\"{lastPageContent}\"");
                }
                catch (Exception ex)
                {
                    // Write Exception into exceptions.log
                    LogException(logFilePathShow, file, ex.Message);
                    // Continue to the next iteration
                    continue;
                }

                finally
                {
                    // Always close the document properly
                    if (doc != null)
                    {
                        doc.Close();
                        Marshal.ReleaseComObject(doc);
                        doc = null;
                    }
                }
            }

            // Quit the Word application
            wordApp.Quit();
            Marshal.ReleaseComObject(wordApp);
            wordApp = null;

            // Clear string array
            DOCXfiles = null;
            Cursor = Cursors.Default;
            TxtBoxLoad.Text = "Chose your folder location ...";
            labelInfo.Text = "Done.";
        }

        private void btnCenterPara_Click(object sender, EventArgs e)
        {
            if (DOCXfiles == null || string.IsNullOrEmpty(TxtBoxLoad.Text))
            {
                labelErrorMessage.Text = "No source folder was selected, Please select one.";
                return;
            }

            else if (DOCXfiles.Length == 0)
            {
                labelErrorMessage.Text = "No DOCX file was found in the selected folder";
                return;
            }

            labelErrorMessage.Text = "";
            Cursor = Cursors.WaitCursor;
            labelInfo.Text = "Processing ...";
            labelErrorMessage.Text = "";

            // Create a new instance of Microsoft Word through the Interop library
            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;

            // Log file
            string logFilePath = selected_path + @"\exceptions.log";

            // Delete the log file if it exists
            if (File.Exists(logFilePath))
            {
                File.Delete(logFilePath);
            }

            foreach (string file in DOCXfiles)
            {
                try
                {
                    wordApp.Visible = false;
                    object missing = System.Reflection.Missing.Value;
                    // Open the document 
                    doc = wordApp.Documents.Open(file);

                    string keyword = txtBoxkeyword.Text;

                    // Center all paragraphs
                    foreach (Paragraph para in doc.Paragraphs)
                    {
                        string text = para.Range.Text.Trim();

                        if (text.Contains(keyword))
                        {
                            para.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            break; // Remove this if you want to center *all* matching paragraphs
                        }
                    }

                    // Save and close
                    doc.Save();
                }
                catch (Exception ex)
                {
                    // Write Exception into exceptions.log
                    LogException(logFilePath, file, ex.Message);
                    // Continue to the next iteration
                    continue;
                }

                finally
                {
                    // Always close the document properly
                    if (doc != null)
                    {
                        doc.Close();
                        Marshal.ReleaseComObject(doc);
                        doc = null;
                    }
                }
            }

            // Quit the Word application
            wordApp.Quit();
            Marshal.ReleaseComObject(wordApp);
            wordApp = null;

            // Clear string array
            DOCXfiles = null;
            Cursor = Cursors.Default;
            TxtBoxLoad.Text = "Chose your folder location ...";
            labelInfo.Text = "Done.";
        }

        private void btnCenterNewline_Click(object sender, EventArgs e)
        {
            if (DOCXfiles == null || string.IsNullOrEmpty(TxtBoxLoad.Text))
            {
                labelErrorMessage.Text = "No source folder was selected, Please select one.";
                return;
            }

            else if (DOCXfiles.Length == 0)
            {
                labelErrorMessage.Text = "No DOCX file was found in the selected folder";
                return;
            }

            labelErrorMessage.Text = "";
            Cursor = Cursors.WaitCursor;
            labelInfo.Text = "Processing ...";
            labelErrorMessage.Text = "";

            // Create a new instance of Microsoft Word through the Interop library
            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;

            // Log file
            string logFilePath = selected_path + @"\exceptions.log";
            // Delete the log file if it exists
            if (File.Exists(logFilePath))
            {
                File.Delete(logFilePath);
            }

            foreach (string file in DOCXfiles)
            {
                try
                {
                    wordApp.Visible = false;
                    object missing = System.Reflection.Missing.Value;
                    // Open the document 
                    doc = wordApp.Documents.Open(file);

                    // Loop through all content controls in the document
                    foreach (ContentControl contentControl in doc.ContentControls)
                    {
                        contentControl.Delete(); // Removes the content control
                    }

                    // Track if replacements occurred
                    // Replace text and count replacements
                    var (replacementsMade, totalOccurrences, textExists) = FindAndReplace(wordApp, txtBoxOld.Text, txtBoxNew.Text);

                    var (replacementsMade2, totalOccurrences2, textExists2) = FindAndReplace(wordApp, txtBoxOld2.Text, txtBoxNew2.Text);

                    var (replacementsMade3, totalOccurrences3, textExists3) = FindAndReplace(wordApp, txtBoxOld3.Text, txtBoxNew3.Text);

                    string keyword = txtBoxkeyword.Text;

                    // Center all paragraphs
                    foreach (Paragraph para in doc.Paragraphs)
                    {
                        string text = para.Range.Text.Trim();

                        if (text.Contains(keyword))
                        {
                            para.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            break; // Remove this if you want to center *all* matching paragraphs
                        }
                    }

                    // Save and close
                    doc.Save();

                    // Log the results with replacement count
                    if (replacementsMade > 0 || replacementsMade2 > 0 || replacementsMade3 > 0 || textExists == true || textExists2 == true || textExists3 == true)
                    {
                        LogException(logFilePath, file, "1er : " + replacementsMade + " Replacement was successful.");
                        LogException(logFilePath, file, "2eme : " + replacementsMade2 + " Replacement was successful.");
                        LogException(logFilePath, file, "3eme : " + replacementsMade3 + " Replacement was successful.");
                    }
                    else
                    {
                        LogException(logFilePath, file, "1er : " + replacementsMade + " No replacements were made.");
                        LogException(logFilePath, file, "2eme : " + replacementsMade2 + " No replacements were made.");
                        LogException(logFilePath, file, "3eme : " + replacementsMade3 + " No replacements were made.");
                    }
                }
                catch (Exception ex)
                {
                    // Write Exception into exceptions.log
                    LogException(logFilePath, file, ex.Message);
                    // Continue to the next iteration
                    continue;
                }

                finally
                {
                    // Always close the document properly
                    if (doc != null)
                    {
                        doc.Close();
                        Marshal.ReleaseComObject(doc);
                        doc = null;
                    }
                }
            }

            // Quit the Word application
            wordApp.Quit();
            Marshal.ReleaseComObject(wordApp);
            wordApp = null;

            // Clear string array
            DOCXfiles = null;
            Cursor = Cursors.Default;
            TxtBoxLoad.Text = "Chose your folder location ...";
            labelInfo.Text = "Done.";
        }

        private void btnReplaceSurface_Click(object sender, EventArgs e)
        {
            if (DOCXfiles == null || string.IsNullOrEmpty(TxtBoxLoad.Text))
            {
                labelErrorMessage.Text = "No source folder was selected, Please select one.";
                return;
            }

            else if (DOCXfiles.Length == 0)
            {
                labelErrorMessage.Text = "No DOCX file was found in the selected folder";
                return;
            }

            labelErrorMessage.Text = "";
            Cursor = Cursors.WaitCursor;
            labelInfo.Text = "Processing ...";
            labelErrorMessage.Text = "";

            // Create a new instance of Microsoft Word through the Interop library
            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;

            // Log file
            string logFilePath = selected_path + @"\exceptions.log";
            // Delete the log file if it exists
            if (File.Exists(logFilePath))
            {
                File.Delete(logFilePath);
            }

            foreach (string file in DOCXfiles)
            {
                try
                {
                    wordApp.Visible = false;
                    object missing = System.Reflection.Missing.Value;
                    // Open the document 
                    doc = wordApp.Documents.Open(file);

                    // Loop through all content controls in the document
                    foreach (ContentControl contentControl in doc.ContentControls)
                    {
                        contentControl.Delete(); // Removes the content control
                    }

                    // Get the total number of pages in the document
                    object what = WdGoToItem.wdGoToPage;
                    object which = WdGoToDirection.wdGoToLast;
                    Range lastPageRange = wordApp.Selection.GoTo(ref what, ref which);

                    int lastPageNumber = lastPageRange.Information[WdInformation.wdActiveEndPageNumber];

                    string lastPageContent = "";

                    foreach (Paragraph para in doc.Paragraphs)
                    {
                        Range paraRange = para.Range;
                        lastPageContent += paraRange.Text;
                    }

                    // Start First Regex to find Surface
                    string pattern = @"\(\s*(?:A|CA|H)[A-Z0-9]*\)";
                    Match match = Regex.Match(lastPageContent, pattern);
                    if (match.Success)
                    {
                        string findText = match.Value;

                        //Detect H an A, CA
                        string clean = findText.Replace(" ", "").Trim('(', ')');  // → "A03CA59H0"

                        // Start Second Regex to detect number
                        string pattern2 = @"A(\d+?)CA(\d+?)H(\d+)";
                        Match match2 = Regex.Match(clean, pattern2);

                        if (match2.Success)
                        {
                            string part1 = match2.Groups[1].Value; // Between A and CA
                            string part2 = match2.Groups[2].Value; // Between CA and H
                            string part3 = match2.Groups[3].Value; // After H

                            // Get the Find object from Word
                            Word.Find findObject = doc.Content.Find;
                            findObject.ClearFormatting();
                            findObject.Text = findText;
                            findObject.Replacement.ClearFormatting();

                            findObject.Replacement.Text =  "( QQ" + part3 + " H " + part2 + " A " + part1 + " CA " + ") ";

                            findObject.MatchCase = true; // For case-sensitive search
                            findObject.MatchWholeWord = true; // For matching whole words only

                            //Set to replace all occurrences
                            object replaceAll = Word.WdReplace.wdReplaceAll;
                            findObject.Execute(Replace: ref replaceAll);
                            // Save and close
                            doc.Save();

                            LogException(logFilePath, file, " Replacement was successful.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Write Exception into exceptions.log
                    LogException(logFilePath, file, ex.Message);
                    // Continue to the next iteration
                    continue;
                }

                finally
                {
                    // Always close the document properly
                    if (doc != null)
                    {
                        doc.Close();
                        Marshal.ReleaseComObject(doc);
                        doc = null;
                    }
                }
            }

            // Quit the Word application
            wordApp.Quit();
            Marshal.ReleaseComObject(wordApp);
            wordApp = null;

            // Clear string array
            DOCXfiles = null;
            Cursor = Cursors.Default;
            TxtBoxLoad.Text = "Chose your folder location ...";
            labelInfo.Text = "Done.";
        }
    }
}
