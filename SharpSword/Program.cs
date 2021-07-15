using System;
using System.IO.Compression;
using System.IO;
using System.Xml;

namespace SharpSword
{
    class Program
    {
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
"Developed By: @sadpanda_sec \n\n" +
"Description: Read Contents of Word Documents (Docx).\n\n" +
"Usage: SharpSword.exe C:\\Some\\Path\\To\\Document.docx");
                System.Environment.Exit(0);

            }
            else if (args.Length == 1)
            {
                if (File.Exists(args[0]) && Path.GetExtension(args[0]).ToLower() == ".docx")
                {
                    var docPath = Path.GetFullPath(args[0]);

                    using (var archive = ZipFile.OpenRead(docPath))
                    {
                     
                        var xmlFile = archive.GetEntry(@"word/document.xml");
                        if (xmlFile == null)
                            return;

                        using (var stream = xmlFile.Open())
                        {
                            using (var reader = new StreamReader(stream))
                            {
                                XmlDocument xmldoc = new XmlDocument();
                                xmldoc.Load(stream);
                                XmlNodeList plaintext = xmldoc.GetElementsByTagName("w:t");
                                DateTime date = DateTime.Now;
                                Console.WriteLine("\n" + date + ": " + "Reading Document: " + docPath + "\n\n");
                                for (int i=0; i < plaintext.Count; i++)
                                {
                                    Console.WriteLine(plaintext[i].InnerText);
                                }
                                System.Environment.Exit(0);
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("File Does Not Exist Or File Extention is Not DOCX");
                    System.Environment.Exit(0);

                }
            }
            else
            {
                if (args.Length > 1)
                {
                    Console.WriteLine("Error...Provided more than one command line argument\n\n" +
                        "Usage: SharpSword.exe C:\\Some\\Path\\To\\Document.docx");
                    System.Environment.Exit(0);
                }
            }
        }

    }
        
}


