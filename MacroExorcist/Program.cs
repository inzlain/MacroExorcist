using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenMcdf;
using System.IO.Compression;
using System.IO;

namespace MacroExorcist
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                string documentFilename = args[0];
                string oleFilename = "";

                Console.WriteLine(String.Format("[+] Saving backup to: {0}.bak", documentFilename));
                if (File.Exists(documentFilename + ".bak")) { File.Delete(documentFilename + ".bak"); }
                File.Copy(documentFilename, documentFilename + ".bak");

                Console.WriteLine("[+] Opening document...");
                try
                {
                    string tempDirectory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
                    Directory.CreateDirectory(tempDirectory);

                    ZipFile.ExtractToDirectory(documentFilename, tempDirectory);
                    if (File.Exists(Path.Combine(tempDirectory, "word", "vbaProject.bin")))
                    {
                        Console.WriteLine("[.] Word document recognised");
                        oleFilename = Path.Combine(tempDirectory, "word", "vbaProject.bin");
                        try
                        {
                            Console.WriteLine("[+] Processing document...");
                            CompoundFile compoundFile = new CompoundFile(oleFilename, CFSUpdateMode.Update, 0);
                            CFStorage commonStorage = compoundFile.RootStorage;

                            Console.WriteLine("[+] Extracting VBA Project...");
                            byte[] vbaProjectStream = commonStorage.GetStorage("VBA").GetStream("_VBA_PROJECT").GetData();

                            string vbaVersion = vbaProjectStream[2].ToString("X2") + vbaProjectStream[3].ToString("X2");
                            if (vbaVersion == "BEEF") {
                                Console.WriteLine(String.Format("[.] VBA project version is: {0}", vbaVersion));
                                Console.WriteLine("[!] Document is already exorcised");
                                return;
                            }
                            else if (vbaVersion == "B200") { vbaVersion += " - Office 2016/2019 (x64)"; }
                            else if (vbaVersion == "B200") { vbaVersion += " - Office 2016/2019 (x64)"; }
                            else if (vbaVersion == "AF00") { vbaVersion += " - Office 2016/2019 (x86)"; }
                            else if (vbaVersion == "A600") { vbaVersion += " - Office 2013 (x64)"; }
                            else if (vbaVersion == "A300") { vbaVersion += " - Office 2013 (x86)"; }
                            else if (vbaVersion == "9700") { vbaVersion += " - Office 2010 (x86)"; }
                            else { vbaVersion += " - Unknown Office Version"; }

                            Console.WriteLine(String.Format("[.] VBA project version is : {0}", vbaVersion));

                            Console.WriteLine("[+] Using BEEF to exorcise VBA demons... \\m/");
                            vbaProjectStream[2] = 0xBE;
                            vbaProjectStream[3] = 0xEF;                                                    
                            commonStorage.GetStorage("VBA").GetStream("_VBA_PROJECT").SetData(vbaProjectStream);

                            compoundFile.Commit();
                            compoundFile.Close();

                            Console.WriteLine("[+] Saving document...");
                            if (File.Exists(documentFilename)) { File.Delete(documentFilename); }
                            ZipFile.CreateFromDirectory(tempDirectory, documentFilename);
                            Directory.Delete(tempDirectory, true);

                            Console.WriteLine("[+] Exorcism successful!");
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(String.Format("[!] Error processing document: {0}", e.ToString()));
                        }
                    }
                    else
                    {
                        Console.WriteLine("[!] Error: document is not a Word document");
                    }

                }
                catch (Exception)
                {
                    Console.WriteLine("[!] Error: document format not recognised");
                }

            }
        }
    }
}
