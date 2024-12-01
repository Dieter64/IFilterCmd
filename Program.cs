// MIT License
//
// IFilterCmd (a simple command line caller for IFilterTextReader by Kees van Spelde <sicos2002@hotmail.com> Magic-Sessions. (www.magic-sessions.com)
//
// Copyright(c) 2022
// Dieter Riekert
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

using System;
using System.IO;
using IFilterTextReader;

/* From IFilterTextReader documentation
/// <summary>
/// When set to <c>true</c> the <see cref="NativeMethods.IFilter"/>
/// doesn't read embedded content, e.g. an attachment inside an E-mail msg file. This parameter is default set to <c>false</c>
/// </summary>
public bool DisableEmbeddedContent { get; set; }

/// <summary>
/// When set to <c>true</c> the metadata properties of
/// a document are also returned, e.g. the summary properties of a Word document. This parameter
/// is default set to <c>false</c>
/// </summary>
public bool IncludeProperties { get; set; }

/// <summary>
/// When set to <c>true</c> the file to process is completely read 
/// into memory first before the iFilters starts to read chunks, when set to <c>false</c> the iFilter reads
/// directly from the file and advances reading when the chunks are returned. 
/// Default set to <c>false</c>
/// </summary>
public bool ReadIntoMemory { get; set; }

/// <summary>
/// Can be used to timeout when parsing very large files, default set to <see cref="FilterReaderTimeout.NoTimeout"/>
/// </summary>
public FilterReaderTimeout ReaderTimeout { get; set; }

/// <summary>
/// The timeout in millisecond when the <see cref="FilterReaderTimeout"/> is set to a value other then <see cref="FilterReaderTimeout.NoTimeout"/>
/// </summary>
/// <remarks>This value is only
/// used when <see cref="FilterReaderTimeout"/> is set to <see cref="FilterReaderTimeout.TimeoutOnly"/>
/// or <see cref="FilterReaderTimeout.TimeoutWithException"/>
/// </remarks>
public int Timeout { get; set; }

/// <summary>
/// Indicates when <c>true</c> (default) that certain characters should be translated to likely ASCII characters.
/// </summary>
public bool DoCleanUpCharacters { get; set; }

/// <summary>
/// The separator that is used between word breaks
/// </summary>
public string WordBreakSeparator { get; set; }
*/
namespace IFilterCmd
{
    class Program
    {
        private static void Usage()
        {
            Console.WriteLine("Usage: IFilterCmd <input file> [-o <output file> -c[+/-] -e[+/-] -p[+/-] -m[+/-] -te <timeout> -ti <timeout> -?]\r\n\r\n" +
                              "IFilterCmd is a simple command line front end for IFilterTextReader available at github.\r\n" +
                              "Options ('+' activates an option (default if missing), '-' deactivates it')\r\n" +
                              "<input file>     : Source file to run IFilter on\r\n" +
                              "-o <output file> : Destination for output. If not provided output is written to the console.\r\n" +
                              "-M               : Multiple files (Creates a txt file for each input file which is a simple pattern).\r\n" +
                              "-e               : Doesn't read embedded content, e.g. an attachment inside an E-mail msg file (default false)\r\n" +
                              "-p               : The metadata properties of a document are also returned (default false)\r\n\r\n" +
                              "less important options:\r\n" +
                              "-c               : Do cleanup characters (default true). Indicates that certain characters should be translated to likely ASCII characters\r\n" +
                              "-m               : Read input file completely into memory before acting (default false)\r\n" +
                              "-w <char>        : Character which identifies a word break. Default is the slash '-'.\r\n" +
                              "-te <timeout>    : Define timeout in millisecons for large files with aborting after timeout elapsed\r\n" +
                              "-ti <timeout>    : Define timeout in millisecons for large files with continuing after timeout elapsed\r\n\r\n" +
                              "Example: IFilterCmd c:\\temp\\test.pdf -o c:\\temp\\output.txt -c- -e- -m+ -ti 5000\r\n\r\n" +
                              "Exit code is 0 for success (you have to still check for an empty output file) and 1 for errors");

            Environment.Exit(1);
        }

        private static void AdditionalArgumentRequired(String arg)
        {
            Console.WriteLine($"Expecting another argument after option {arg}");
            Usage();
        }

        static void Main(string[] args)
        {
            String inputFile = null;
            String outputFile = null;
            bool multipleFiles = false;

            var options = new FilterReaderOptions();

            for (int i = 0; i < args.Length; i++)
            {
                String arg = args[i];

                if (arg.StartsWith("-") || arg.StartsWith("/"))
                {
                    if (arg.Length < 2)
                    {
                        Console.WriteLine($"Option missing after - or /");
                        Usage();
                    }

                    switch (arg.Substring(1))
                    {
                        case "M":
                            multipleFiles = true;
                            break;
                        case "c":
                        case "c+":
                            // Default is true. 
                            options.DoCleanUpCharacters = true;
                            break;
                        case "c-":
                            options.DoCleanUpCharacters = false;
                            break;
                        case "e":
                        case "e+":
                            // Default is false
                            options.DisableEmbeddedContent = true;
                            break;
                        case "e-":
                            // Default is false
                            options.DisableEmbeddedContent = false;
                            break;
                        case "p":
                        case "p+":
                            // Default false
                            options.IncludeProperties = true;
                            break;
                        case "p-":
                            options.IncludeProperties = false;
                            break;
                        case "m":
                        case "m+":
                            // default false
                            options.ReadIntoMemory = true;
                            break;
                        case "m-":
                            options.ReadIntoMemory = false;
                            break;
                        case "te":
                        case "ti":
                            if (i < args.Length - 1)
                            {
                                String timeoutValue = args[i + 1];
                                if (!Int32.TryParse(timeoutValue, out int timeout))
                                {
                                    Console.WriteLine($"Invalid timeout value {timeoutValue}");
                                    Usage();
                                }
                                options.Timeout = timeout;
                                options.ReaderTimeout = arg.Substring(1, 1) == "i" ? FilterReaderTimeout.TimeoutOnly : FilterReaderTimeout.TimeoutWithException;
                                i++;
                            }
                            else
                            {
                                AdditionalArgumentRequired(arg);
                            }
                            break;
                        case "o":
                            if (i < args.Length - 1)
                            {
                                outputFile = args[i + 1];
                                i++;
                            }
                            else
                            {
                                AdditionalArgumentRequired(arg);
                            }
                            break;
                        case "w":
                            // Seems not to work. Omit in Usage()
                            if (i < args.Length - 1)
                            {
                                options.WordBreakSeparator = args[i + 1];
                                i++;
                            }
                            else
                            {
                                AdditionalArgumentRequired(arg);
                            }
                            break;
                        case "?":
                            Usage();
                            Environment.Exit(1);
                            return;
                        default:
                            Console.WriteLine($"Unknown option {arg}");
                            Usage();
                            break;
                    }
                }
                else
                {
                    if (!String.IsNullOrEmpty(inputFile))
                    {
                        Console.WriteLine($"Only one input file allowed! Found more than one. First one:{inputFile}, next:{arg}");
                        Usage();
                    }
                    inputFile = arg;
                }
            }

            if (String.IsNullOrEmpty(inputFile))
            {
                Console.WriteLine($"No input file provided");
                Usage();
            }

            if (multipleFiles)
            {
                string[] fileNames = Directory.GetFiles(".", inputFile);
                foreach (var file in fileNames)
                {
                    try
                    {
                        string outFile = file + ".txt";
                        using (var reader = new FilterReader(file, string.Empty, options))
                        {
                            string line;

                            using (TextWriter outStream = File.CreateText(outFile))
                            {
                                while ((line = reader.ReadLine()) != null)
                                {
                                    outStream.WriteLine(line);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Exception call in IFIlterCmd for file:{file}: {ex.Message}");
                        // Environment.Exit(1);
                    }
                }
            }
            else
            {
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"The input file '{inputFile}' does not exists");
                    Usage();
                }

                try
                {
                    using (var reader = new FilterReader(inputFile, string.Empty, options))
                    {
                        string line;

                        using (TextWriter outStream = String.IsNullOrEmpty(outputFile) ? Console.Out : File.CreateText(outputFile))
                        {
                            while ((line = reader.ReadLine()) != null)
                            {
                                outStream.WriteLine(line);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Exception call in IFIlterCmd:{ex.Message}");
                    Environment.Exit(1);
                }
            }
        }
    }
}
