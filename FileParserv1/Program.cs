using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;


namespace FileParserv1
{


    class Program
    {
        static void Main()
        {
            string filePath = "C:\\Users\\GentritSelimi\\Desktop\\EA1024PR(39).txt";

            string[] lines = File.ReadAllLines(filePath);

            // Dictionary to store headers and their values
            var seq1Values = new Dictionary<string, string>();
            var generalInfo = new Dictionary<string, string>();

            var isFirstSequence = true;
            bool isSecSeq, isThirdSeq, isFourthSeq, isFifthSeq, isSixSeq, isSevenSeq, isEightSeq, isNineSeq, isTenSeq, isElevenSeq;
            isSecSeq = isThirdSeq = isFourthSeq = isFifthSeq = isSixSeq = isSevenSeq = isEightSeq = isNineSeq = isTenSeq = isElevenSeq = false;

            var serialNumbers = new List<string>();
            var modelNames = new List<string>();
            var timeStamps = new List<string>();

            var data = new Dictionary<string, List<string>>();
            var testStatus = new Dictionary<string, string>();

            for (int k = 0; k < lines.Length; k++)
            {
                string cleanLine = lines[k].Trim();

                if (cleanLine.Contains("Model") && lines[k + 3].Contains("YYYY_MM_DD"))
                {
                    // Extract and store the values
                    generalInfo["Model Name"] = ExtractValue(cleanLine, "Model Name:");
                    generalInfo["Customer"] = ExtractValue(cleanLine, "Customer:");
                    generalInfo["Serial No"] = ExtractValue(cleanLine, "Serial No:");
                    generalInfo["Order No"] = ExtractValue(lines[k + 1], "Order No.:");
                    generalInfo["Lot No"] = ExtractValue(lines[k + 1], "Lot No.:");
                    generalInfo["Total Load No"] = ExtractValue(lines[k + 1], "Total Load No.:");
                    generalInfo["Environment"] = ExtractValue(lines[k + 2], "Environment:");
                    generalInfo["Inspector"] = ExtractValue(lines[k + 2], "Inspector:");
                    generalInfo["YYYY_MM_DD"] = ExtractValue(lines[k + 3], "YYYY_MM_DD:");
                    generalInfo["Begin Time"] = ExtractValue(lines[k + 3], "Begin Time:");
                    generalInfo["End Time"] = ExtractValue(lines[k + 3], "End Time:");

                    serialNumbers.Add(generalInfo["Serial No"]);
                    modelNames.Add(generalInfo["Model Name"]);
                    timeStamps.Add(generalInfo["YYYY_MM_DD"]);
                    
                }
            }
            var testPassed = false;
            var sn = string.Empty;
            for (int j = 0; j < lines.Length; j++)
            {
                string cleanLine = lines[j].Trim();

                if (cleanLine.Contains("Model") && lines[j + 3].Contains("YYYY_MM_DD"))
                {
                    var serialNumber = ExtractValue(cleanLine, "Serial No:");
                    sn = serialNumber;
                    data.Add(serialNumber, new List<string>());
                    testStatus.Add(serialNumber, "");
                    seq1Values.Clear();
                    testPassed = false;

                }

                if (cleanLine.StartsWith("SEQ.1:") && !cleanLine.StartsWith("SEQ.11:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isFirstSequence = true;
                    isElevenSeq = false;
                }
                else if (cleanLine.Contains("SEQ.2:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isFirstSequence = false;
                    isSecSeq = true;
                }

                else if (cleanLine.Contains("SEQ.3:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isSecSeq = false;
                    isThirdSeq = true;
                }
                else if (cleanLine.Contains("SEQ.4:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isThirdSeq = false;
                    isFourthSeq = true;
                }
                else if (cleanLine.Contains("SEQ.5:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isFourthSeq = false;
                    isFifthSeq = true;
                }

                else if(cleanLine.Contains("SEQ.6:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isSixSeq = true;
                    isFifthSeq = false;
                }
                else if(cleanLine.Contains("SEQ.7:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isSevenSeq = true;
                    isSixSeq = false;
                }

                else if (cleanLine.Contains("SEQ.8:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isSevenSeq = false;
                    isEightSeq = true;
                }

                else if(cleanLine.Contains("SEQ.9:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isEightSeq = false;
                    isNineSeq = true;
                }
                else if(cleanLine.Contains("SEQ.10:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isNineSeq = false;
                    isTenSeq = true;
                }
                else if (cleanLine.StartsWith("SEQ.11:") && !cleanLine.StartsWith("SEQ.1:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isTenSeq = false;
                    isElevenSeq = true;
                    isFirstSequence = false;
                }

                if (isFirstSequence)
                {
                    // Extract key-value pairs using regex or string split for '=' sign
                    if (cleanLine.Contains("="))
                    {
                        // Example: Vin Port       =        1      Vin Type        =       AC
                        string[] parts = Regex.Split(cleanLine, @"\s+=\s+|\s{2,}");

                        for (int i = 0; i < parts.Length - 1; i += 2)
                        {
                            string key = parts[i];
                            string value = parts[i + 1];

                            // Store in the dictionary or process further
                            seq1Values["Seq1-" + key] = value;
                        }
                    }

                    if (cleanLine.Contains("Load Name") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").ToList();
                        values.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq1-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Ld    TRIG") && lines[j + 2].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 2], @"\s{2,}").ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);

                        values = values.Where(x => !string.IsNullOrEmpty(x)).ToList();

                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq1-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Ld    Ton") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").ToList();

                        values.RemoveAt(0);
                        keys.RemoveAt(0);

                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq1-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    //TBD tomorrow.
                    if (cleanLine.Contains("Ref Ton from LOAD:") && lines[j + 5].Contains("Tdls"))
                    {
                        var keys = Regex.Split(lines[j + 1], @"\s{2,}").ToList();


                    }
                }
                if (isSecSeq || isThirdSeq || isFourthSeq)
                {
                    var seqValue = isSecSeq ? "Seq2-" : isThirdSeq ? "Seq3-" : "Seq4-";
                    if (cleanLine.Contains('='))
                    {
                        // Example: Vin Port       =        1      Vin Type        =       AC
                        string[] parts = Regex.Split(cleanLine, @"\s+=\s+|\s{2,}");

                        for (int i = 0; i < parts.Length - 1; i += 2)
                        {
                            string key = parts[i];
                            string value = parts[i + 1];

                            // Store in the dictionary or process further
                            seq1Values[$"{seqValue}" + key] = value;
                        }
                    }
                    if (cleanLine.Contains("Load Name") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").ToList();
                        values.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values[$"{seqValue}" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("BITS-1") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"(?<!SLEW)\s{1,}(?!Rate)").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").ToList();

                        values.RemoveAt(0);
                        keys.RemoveAt(0);

                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values[$"{seqValue}" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if ((cleanLine.Contains("Vpp Max") || cleanLine.Contains("Vdc Max")) && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();

                        values.RemoveAt(0);
                        keys.RemoveAt(0);

                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values[$"{seqValue}" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("dV(+) Max") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();

                        values.RemoveAt(0);
                        keys.RemoveAt(0);

                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values[$"{seqValue}" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Vn Max") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();

                        values.RemoveAt(0);
                        keys.RemoveAt(0);

                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values[$"{seqValue}" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }
                }
                if (isFifthSeq)
                {
                    if (cleanLine.Contains('='))
                    {
                        // Example: Vin Port       =        1      Vin Type        =       AC
                        string[] parts = Regex.Split(cleanLine, @"\s+=\s+|\s{2,}");

                        for (int i = 0; i < parts.Length - 1; i += 2)
                        {
                            string key = parts[i];
                            string value = parts[i + 1];

                            // Store in the dictionary or process further
                            seq1Values[$"Seq5-" + key] = value;
                        }
                    }
                    if (cleanLine.Contains("Load Name") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq5-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Period-1") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq5-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("I/R-1") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq5-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Vs Max") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq5-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }
                }

                if (isSixSeq)
                {
                    if (cleanLine.Contains('='))
                    {
                        // Example: Vin Port       =        1      Vin Type        =       AC
                        string[] parts = Regex.Split(cleanLine, @"\s+=\s+|\s{2,}");

                        for (int i = 0; i < parts.Length - 1; i += 2)
                        {
                            string key = parts[i];
                            string value = parts[i + 1];

                            // Store in the dictionary or process further
                            seq1Values[$"Seq6-" + key] = value;
                        }
                    }

                    if (cleanLine.Contains("Load Name") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq6-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }
                }

                if (isSevenSeq)
                {
                    if (cleanLine.Contains('='))
                    {
                        // Example: Vin Port       =        1      Vin Type        =       AC
                        string[] parts = Regex.Split(cleanLine, @"\s+=\s+|\s{2,}");

                        for (int i = 0; i < parts.Length - 1; i += 2)
                        {
                            string key = parts[i];
                            string value = parts[i + 1];

                            // Store in the dictionary or process further
                            seq1Values[$"Seq7-" + key] = value;
                        }
                    }

                    if (cleanLine.Contains("Load Name") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq7-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Vdisable Max") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq7-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }
                }

                if (isEightSeq || isNineSeq)
                {

                    var seqValue = isEightSeq ? "Seq8-" : "Seq9-";
                    if (cleanLine.Contains('='))
                    {
                        // Example: Vin Port       =        1      Vin Type        =       AC
                        string[] parts = Regex.Split(cleanLine, @"\s+=\s+|\s{2,}");

                        for (int i = 0; i < parts.Length - 1; i += 2)
                        {
                            string key = parts[i];
                            string value = parts[i + 1];

                            // Store in the dictionary or process further
                            seq1Values[$"{seqValue}" + key] = value;
                        }
                    }

                    if (cleanLine.Contains("Load Name") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values[$"{seqValue}" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("RISE") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values[$"{seqValue}" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Vdc Max") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values[$"{seqValue}" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }
                }
                if (isTenSeq)
                {
                    if (cleanLine.Contains('='))
                    {
                        // Example: Vin Port       =        1      Vin Type        =       AC
                        string[] parts = Regex.Split(cleanLine, @"\s+=\s+|\s{2,}");

                        for (int i = 0; i < parts.Length - 1; i += 2)
                        {
                            string key = parts[i];
                            string value = parts[i + 1];

                            // Store in the dictionary or process further
                            seq1Values[$"Seq10-" + key] = value;
                        }
                    }

                    if (cleanLine.Contains("Load Name") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq10-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("TRIG") && lines[j + 2].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        keys = keys.Select((item, index) => item.Contains("TRIGG") ? $"TRIGG{index + 1}" : item).ToList();
                        var values = Regex.Split(lines[j + 2], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);

                        values = values.Where(x => !string.IsNullOrEmpty(x)).ToList();

                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq10-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Ton Max") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq10-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }
                    if (cleanLine.Contains("Tons Source") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq10-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }
                }
                if (isElevenSeq)
                {
                    if (cleanLine.Contains('='))
                    {
                        // Example: Vin Port       =        1      Vin Type        =       AC
                        string[] parts = Regex.Split(cleanLine, @"\s+=\s+|\s{2,}");

                        for (int i = 0; i < parts.Length - 1; i += 2)
                        {
                            string key = parts[i];
                            string value = parts[i + 1];

                            // Store in the dictionary or process further
                            seq1Values[$"Seq11-" + key] = value;
                        }
                    }

                    if (cleanLine.Contains("Load Name") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq11-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("TRIG") && lines[j + 2].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        keys = keys.Select((item, index) => item.Contains("TRIGG") ? $"TRIGG{index + 1}" : item).ToList();
                        var values = Regex.Split(lines[j + 2], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);

                        values = values.Where(x => !string.IsNullOrEmpty(x)).ToList();

                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq11-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Thd Max") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq11-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("-----------------------------------------------------------------------"))
                    {
                        data[sn] = seq1Values.Values.ToList();
                        testStatus[sn] = testPassed ? "PASS" : "FALSE";
                    }
                }
                
            }

            var headers = seq1Values.Keys.ToList();
            WriteToCsv(serialNumbers, timeStamps, data, headers,testStatus);
        }

        public static void WriteToCsv(List<string> serialNumbers, List<string>  timeStamps, Dictionary<string,List<string>> data, List<string> headers, 
            Dictionary<string,string> testStatuses)
        {
            var config = new CsvConfiguration(cultureInfo: CultureInfo.InvariantCulture)
            {
                Delimiter = ";",
                HasHeaderRecord = false
            };

            using var writer = new StreamWriter("C:\\Users\\GentritSelimi\\Desktop\\adatper-edac1.csv");
            using var csv = new CsvWriter(writer, config);
            // Skip to row 3 (as row 1 and 2 are empty)
            for (int i = 0; i < 3; i++) // Start counting from 0, so skip 2 rows for row 3
            {
                csv.NextRecord();
            }

            csv.WriteField("SerialNumber");
            csv.WriteField("IsPass");
            csv.WriteField("StartDate");

            foreach (var header in headers)
            {
                csv.WriteField(header);
            }
            csv.NextRecord();

            // Skip to row 4 for writing serial numbers (no more need for row skips since we're already there)
            // Write serial numbers starting from column 1 (row 4 onward)

            for (int i = 0; i < serialNumbers.Count; i++)
            {
                var sn = serialNumbers[i];
                var dictData = data[sn];
                var testStatus = testStatuses[sn];
                var timeStamp = timeStamps[i];

                csv.WriteField(sn);

                for (int j = 0; j < 2; j++)
                {
                    if(j == 0)
                    {
                        csv.WriteField(testStatus);
                    }
                    if(j == 1)
                    {
                        csv.WriteField(timeStamp);
                    }
                }

                foreach (var d in dictData)
                {
                    csv.WriteField(d);
                }
                csv.NextRecord();
            }

            // Now, write dictionary data horizontally starting from row 6, column 4
            for (int i = 0; i < 2; i++) // Skip 2 more rows to get to row 6
            {
                csv.NextRecord();
            }

        }
        static string ExtractValue(string line, string key)
        {
            int startIndex = line.IndexOf(key) + key.Length;
            if (startIndex >= key.Length)
            {
                return line.Substring(startIndex).Trim();
            }
            return string.Empty;
        }

    }
}
