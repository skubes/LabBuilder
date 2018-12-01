using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.IO.Compression;

namespace CreateListStringFromTextFile
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputFile = string.Empty;
            string encodingsetting = string.Empty;
            if (args.Length == 2)
            {

                inputFile = args[0];
                encodingsetting = args[1];
            }
            else
            {
                Console.WriteLine("Must supply two args: path to text file, and encoding. For encoding use 1 for System default(ansi) codepage and 2 for UTF8");
                return;
            }


            StreamReader sr = null;

            // try to open file
            try
            {

                if (encodingsetting == "2")
                {
                    sr = new StreamReader(inputFile, Encoding.UTF8);
                }
                else
                {

                    sr = new StreamReader(inputFile, Encoding.Default);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Exception opening file: {0}", e.Message);
                return;
            }
          

            var inputFilePath = Path.GetFullPath(inputFile);
            Console.WriteLine("Running CreateListStringFromTextFile.exe");
            Console.WriteLine("Opened {0}", inputFilePath);

            var list = new List<string>();
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                list.Add(line);

            }
            
            Console.WriteLine("Added {0} items.", list.Count);
              
            string outputPath = inputFilePath + ".bin";
           

            var binFormatter = new BinaryFormatter();
            FileStream outfile = null;
            try
            {
                outfile = File.Open(outputPath, FileMode.Create);
            
                using (DeflateStream ds = new DeflateStream(outfile, CompressionLevel.Optimal))
                {
                    outfile = null;
                    binFormatter.Serialize(ds, list);
                }
            }
            finally
            {
                if (outfile != null)
                {
                    outfile.Dispose();
                }
                    
            }
           
            
           

            Console.WriteLine("Done.");


        }
    }
}
