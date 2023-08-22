using Microsoft.Office.Interop.Word;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace SharpSword
{
    class Program
    {
        static bool IsProcessRunning(string processName)
        {

            Process[] processes = Process.GetProcessesByName(processName.ToLower());

            foreach (Process process in processes)
            {
                try
                {
                    if (process.SessionId == Process.GetCurrentProcess().SessionId)
                    {
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error : " + ex);
                }
            }

            return false;


        }


        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine(@"

 _____ _                      _____                       _ 
/  ___| |                    /  ___|                     | |
\ `--.| |__   __ _ _ __ _ __ \ `--.__      _____  _ __ __| |
 `--. \ '_ \ / _` | '__| '_ \ `--. \ \ /\ / / _ \| '__/ _` |
/\__/ / | | | (_| | |  | |_) /\__/ /\ V  V / (_) | | | (_| |
\____/|_| |_|\__,_|_|  | .__/\____/  \_/\_/ \___/|_|  \__,_|
                       | |                                  
                       |_|                                  
" +
"" +
"Developed By: @sadpanda_sec & @C0mmand3rOps3c \n\n" +
"Description: Read Contents of Word Documents using MS Office Interop.\n\n" +
"Usage: SharpSword.exe C:\\Some\\Path\\To\\Document.(doc/docm/docx/etc...) [-checkPassword] -[password <password>]\n" +
"Examples:\n" +
"   -SharpSword.exe test.doc : read the contents of a word doc\n" +
"   -SharpSword.exe test.doc -checkPassword : checks if the document is password protected\n" +
"   -SharpSword.exe test.doc -password <somepassword> : decrypts the password protected document and reads contents");
                System.Environment.Exit(0);

            }

            string docFullPath = Path.GetFullPath(args[0]);
            string docName = Path.GetFileName(docFullPath);
            bool checkPassword = false;
            string documentPassword = null;

            for (int i = 1; i < args.Length; i++)
            {
                if (args[i] == "-checkPassword")
                {
                    checkPassword = true;
                }
                else if (args[i] == "-password" && i + 1 < args.Length)
                {
                    documentPassword = args[i + 1];
                }
            }


            if (File.Exists(docFullPath) && Path.GetExtension(docFullPath).Contains("doc"))
            {
                bool isWordRunning = IsProcessRunning("winword");
                Application wordApp = null;
                Document doc = null;
                bool isWordOpen = false;
                bool isDocOpen = false;
                bool isPWprotected = false;

                if (isWordRunning)
                {
                    Console.WriteLine("OPSEC WARNING: Microsoft Word is currently running...Using existing Winword Application\n");
                    wordApp = (Application)Marshal.GetActiveObject("Word.Application");
                    wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                    isWordOpen = true;
                }
                else
                {
                    Console.WriteLine("Microsoft Word is not running...Using New COM Winword Application. \n");
                    wordApp = new Application();
                    wordApp.Visible = false;
                    wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                    isWordOpen = false;
                }


                try
                {

                    if (checkPassword)
                    {

                        if (isWordOpen)
                        {
                            foreach (Document docs in wordApp.Documents)
                            {
                                if (string.Equals(docs.FullName, docFullPath, StringComparison.OrdinalIgnoreCase))
                                {
                                    Console.WriteLine($"OPSEC WARNING: Document '{docName}' is already open by user...\n" +
                                        $"By default, this check will always pass as document is unprotected.\n" +
                                        $"Run this command again when the document is no longer opened by the user.");
                                    isDocOpen = true;
                                    doc = docs;
                                }
                            }

                            if (!isDocOpen)
                            {

                                try
                                {

                                    doc = wordApp.Documents.Open(docFullPath, ReadOnly: true, PasswordDocument: " ", Visible: false);
                                    Console.WriteLine("The document is NOT password protected.");
                                }
                                catch
                                {
                                    Console.WriteLine("WARNING: The document is password protected.");
                                    isPWprotected = true;
                                }

                            }

                        }
                        else
                        {
                            try
                            {
                                doc = wordApp.Documents.Open(docFullPath, ReadOnly: true, PasswordDocument: " ", Visible: false);
                                Console.WriteLine("The document is NOT password protected.");
                            }
                            catch
                            {
                                Console.WriteLine("WARNING: The document is password protected.");
                                isPWprotected = true;
                            }

                        }

                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(documentPassword))
                        {

                            foreach (Document docs in wordApp.Documents)
                            {
                                if (string.Equals(docs.FullName, docFullPath, StringComparison.OrdinalIgnoreCase))
                                {
                                    Console.WriteLine($"Document '{docName}' is already open.");
                                    isDocOpen = true;
                                    doc = docs;
                                }
                            }
                            if (!isDocOpen)
                            {
                                doc = wordApp.Documents.Open(docFullPath, ReadOnly: true, PasswordDocument: documentPassword, Visible: false);

                            }

                            string content = doc.Content.Text;
                            DateTime date = DateTime.Now;
                            Console.WriteLine("\n" + date + ": " + "Reading Document: " + docName + "\n\n");
                            Console.WriteLine("File Content:");
                            Console.WriteLine(content);
                        }

                    }

                    if (!checkPassword && string.IsNullOrEmpty(documentPassword))
                    {
                        try
                        {

                            foreach (Document docs in wordApp.Documents)
                            {

                                if (string.Equals(docs.FullName, docFullPath, StringComparison.OrdinalIgnoreCase))
                                {
                                    Console.WriteLine($"Document '{docName}' is already open.");
                                    isDocOpen = true;
                                    doc = docs;
                                }
                            }

                            if (!isDocOpen)
                            {
                                doc = wordApp.Documents.Open(docFullPath, ReadOnly: true, Visible: false);
                            }

                            string content = doc.Content.Text;
                            DateTime date = DateTime.Now;
                            Console.WriteLine("\n" + date + ": " + "Reading Document: " + docName + "\n\n");
                            Console.WriteLine("File Content:");
                            Console.WriteLine(content);


                        }
                        catch (COMException ex)
                        {
                            Console.WriteLine("Error: " + ex);
                        }


                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex);
                }
                finally
                {

                    if (isDocOpen == false && isPWprotected == false)
                    {

                        doc.Close(false, null, null);

                    }

                    if (isWordOpen == false)
                    {
                        wordApp.Quit(false, null, null);
                        Marshal.ReleaseComObject(wordApp);
                        wordApp = null;


                    }
                    else
                    {
                        wordApp.Visible = true;

                    }
                    if (doc != null)
                    {
                        Marshal.ReleaseComObject(doc);
                        doc = null;
                    }


                }
            }
            else
            {
                Console.WriteLine("File Does Not Exist Or File Extention is Not an MSWord Doc");

            }
        }
    }
}

