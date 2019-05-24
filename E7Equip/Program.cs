using IronOcr;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace E7Equip
{
    class Program
    {
        static string json = "";
        static void Main(string[] args)
        {
            DirectoryInfo dir = new DirectoryInfo(@"C:\Users\Double\Desktop\e7");

            foreach(FileInfo f in dir.EnumerateFiles())
            {
                GetEquipJson(f);
            }
        }

        private static void GetEquipJson(FileInfo f)
        {
            var Ocr = new AdvancedOcr()
            {
                CleanBackgroundNoise = false,
                EnhanceContrast = false,
                EnhanceResolution = false,
                Language = IronOcr.Languages.English.OcrLanguagePack,
                Strategy = IronOcr.AdvancedOcr.OcrStrategy.Advanced,
                ColorSpace = AdvancedOcr.OcrColorSpace.Color,
                DetectWhiteTextOnDarkBackgrounds = true,
                InputImageType = AdvancedOcr.InputTypes.Snippet,
                RotateAndStraighten = false,
                ReadBarCodes = false,
                ColorDepth = 0
            };

            var Results = Ocr.Read(f.FullName);

            var text = Results.Text;


            bool foundSlot = false;
            bool foundRarity = false;
            bool foundSet = false;
            bool foundMain = false;
            bool foundSub1 = false;
            bool foundSub2 = false;
            bool foundSub3 = false;
            bool foundSub4 = false;

            //  "Armor", "Boots", "Helmet", "Necklace", "Ring", "Weapon"
            string[] slots = { "Armor", "Boots", "Helmet", "Necklace", "Ring", "Weapon"};
            string[] rarities = { "Epic", "Heroic", "Rare", "Normal", "Good" };
            string[] stats = { "Defense", "Speed", "Attack", "Health", "Effectiveness", "Effect Resistance", "Critical Hit Chance", "Critial Hit Damage" };

            Item item = new Item();

            foreach (string line in text.Split(new[] { '\r', '\n' }))
            {

                // Slot and Rirty forst
                if(!foundSlot || !foundRarity)
                {
                    var slot = slots.FirstOrDefault<string>(s => line.Contains(s));
                    switch (slot)
                    {
                        case "Armor":
                        case "Boots":
                        case "Helmet":
                        case "Necklace":
                        case "Ring":
                        case "Weapon":
                            item.slot = slot;
                            foundSlot = true;
                            break;
                        default:
                            break;
                    }

                    var rarity = rarities.FirstOrDefault<string>(s => line.Contains(s));
                    switch (rarity)
                    {
                        case "Epic":
                        case "Heroic":
                        case "Rare":
                        case "Good":
                        case "Normal":
                            item.rarity = rarity;
                            foundRarity = true;
                            break;
                        default:
                            break;
                    }

                    continue;
                }

                if(!foundMain)
                {
                    // Look for first main stat
                    var stat = stats.FirstOrDefault<string>(s => line.Contains(s));
                    bool per = false;
                    int val = 0;
                    string reg = @" *(\d+)";
                    Regex rx;
                    MatchCollection matches;
                    switch (stat)
                    {
                        case "Defense":
                            per = stat.Contains("%");
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if(matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            if (per)
                            {
                                item.mainStat.Add("DefP");
                                item.mainStat.Add(val);
                            }
                            else
                            {
                                item.mainStat.Add("Def");
                                item.mainStat.Add(val);
                            }
                            foundMain = true;
                            break;
                        case "Speed":
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.mainStat = new List<dynamic>();
                            item.mainStat.Add("Spd");
                            item.mainStat.Add(val);
                            foundMain = true;
                            break;
                        case "Attack":
                            per = stat.Contains("%");
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.mainStat = new List<dynamic>();
                            if (per)
                            {
                                item.mainStat.Add("AtkP");
                                item.mainStat.Add(val);
                            }
                            else
                            {
                                item.mainStat.Add("Atk");
                                item.mainStat.Add(val);
                            }
                            foundMain = true;
                            break;
                        case "Health":
                            per = stat.Contains("%");
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.mainStat = new List<dynamic>();
                            if (per)
                            {
                                item.mainStat.Add("HPP");
                                item.mainStat.Add(val);
                            }
                            else
                            {
                                item.mainStat.Add("HP");
                                item.mainStat.Add(val);
                            }
                            foundMain = true;
                            break;
                        case "Effectiveness":
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.mainStat = new List<dynamic>();
                            item.mainStat.Add("Eff");
                            item.mainStat.Add(val);
                            foundMain = true;
                            break;
                        case "Effect Resistance":
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.mainStat = new List<dynamic>();
                            item.mainStat.Add("Res");
                            item.mainStat.Add(val);
                            foundMain = true;
                            break;
                        case "Critical Hit Chance":
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.mainStat = new List<dynamic>();
                            item.mainStat.Add("CChance");
                            item.mainStat.Add(val);
                            foundMain = true;
                            break;
                        case "Critial Hit Damage":
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.mainStat = new List<dynamic>();
                            item.mainStat.Add("CDmg");
                            item.mainStat.Add(val);
                            foundMain = true;
                            break;
                        default:
                            //Do work
                            break;
                    }

                    continue;
                }

                if (!foundMain)
                {
                    // Look for first main stat
                    var stat = stats.FirstOrDefault<string>(s => line.Contains(s));
                    bool per = false;
                    int val = 0;
                    string reg = @" *(\d+)";
                    Regex rx;
                    MatchCollection matches;
                    switch (stat)
                    {
                        case "Defense":
                            per = stat.Contains("%");
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            if (per)
                            {
                                item.mainStat.Add("DefP");
                                item.mainStat.Add(val);
                            }
                            else
                            {
                                item.mainStat.Add("Def");
                                item.mainStat.Add(val);
                            }
                            foundMain = true;
                            break;
                        case "Speed":
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.mainStat = new List<dynamic>();
                            item.mainStat.Add("Spd");
                            item.mainStat.Add(val);
                            foundMain = true;
                            break;
                        case "Attack":
                            per = stat.Contains("%");
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.mainStat = new List<dynamic>();
                            if (per)
                            {
                                item.mainStat.Add("AtkP");
                                item.mainStat.Add(val);
                            }
                            else
                            {
                                item.mainStat.Add("Atk");
                                item.mainStat.Add(val);
                            }
                            foundMain = true;
                            break;
                        case "Health":
                            per = stat.Contains("%");
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.mainStat = new List<dynamic>();
                            if (per)
                            {
                                item.mainStat.Add("HPP");
                                item.mainStat.Add(val);
                            }
                            else
                            {
                                item.mainStat.Add("HP");
                                item.mainStat.Add(val);
                            }
                            foundMain = true;
                            break;
                        case "Effectiveness":
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.mainStat = new List<dynamic>();
                            item.mainStat.Add("Eff");
                            item.mainStat.Add(val);
                            foundMain = true;
                            break;
                        case "Effect Resistance":
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.mainStat = new List<dynamic>();
                            item.mainStat.Add("Res");
                            item.mainStat.Add(val);
                            foundMain = true;
                            break;
                        case "Critical Hit Chance":
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.mainStat = new List<dynamic>();
                            item.mainStat.Add("CChance");
                            item.mainStat.Add(val);
                            foundMain = true;
                            break;
                        case "Critial Hit Damage":
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.mainStat = new List<dynamic>();
                            item.mainStat.Add("CDmg");
                            item.mainStat.Add(val);
                            foundMain = true;
                            break;
                        default:
                            //Do work
                            break;
                    }

                    continue;
                }

                if (!foundSub1)
                {
                    // Look for first main stat
                    var stat = stats.FirstOrDefault<string>(s => line.Contains(s));
                    bool per = false;
                    int val = 0;
                    string reg = @" *(\d+)";
                    Regex rx;
                    MatchCollection matches;
                    switch (stat)
                    {
                        case "Defense":
                            per = stat.Contains("%");
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            if (per)
                            {
                                item.subStat1.Add("DefP");
                                item.subStat1.Add(val);
                            }
                            else
                            {
                                item.subStat1.Add("Def");
                                item.subStat1.Add(val);
                            }
                            foundSub1 = true;
                            break;
                        case "Speed":
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.subStat1 = new List<dynamic>();
                            item.subStat1.Add("Spd");
                            item.subStat1.Add(val);
                            foundSub1 = true;
                            break;
                        case "Attack":
                            per = stat.Contains("%");
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.subStat1 = new List<dynamic>();
                            if (per)
                            {
                                item.subStat1.Add("AtkP");
                                item.subStat1.Add(val);
                            }
                            else
                            {
                                item.subStat1.Add("Atk");
                                item.subStat1.Add(val);
                            }
                            foundSub1 = true;
                            break;
                        case "Health":
                            per = stat.Contains("%");
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.subStat1 = new List<dynamic>();
                            if (per)
                            {
                                item.subStat1.Add("HPP");
                                item.subStat1.Add(val);
                            }
                            else
                            {
                                item.subStat1.Add("HP");
                                item.subStat1.Add(val);
                            }
                            foundSub1 = true;
                            break;
                        case "Effectiveness":
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.subStat1 = new List<dynamic>();
                            item.subStat1.Add("Eff");
                            item.subStat1.Add(val);
                            foundSub1 = true;
                            break;
                        case "Effect Resistance":
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.subStat1 = new List<dynamic>();
                            item.subStat1.Add("Res");
                            item.subStat1.Add(val);
                            foundSub1 = true;
                            break;
                        case "Critical Hit Chance":
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.subStat1 = new List<dynamic>();
                            item.subStat1.Add("CChance");
                            item.subStat1.Add(val);
                            foundSub1 = true;
                            break;
                        case "Critial Hit Damage":
                            rx = new Regex(stat + reg, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            matches = rx.Matches(line);
                            if (matches.Count == 0) { break; }
                            val = Int32.Parse(matches[0].Groups[1].Value);
                            item.subStat1 = new List<dynamic>();
                            item.subStat1.Add("CDmg");
                            item.subStat1.Add(val);
                            foundSub1 = true;
                            break;
                        default:
                            //Do work
                            break;
                    }

                    continue;
                }

            }

            string output = JsonConvert.SerializeObject(item);
            Console.WriteLine(output);
        }

        public static List<Item> items = new List<Item>();
    }

    
    public class Item
    {
        public int ability;
        public int level;
        public string set;
        public string slot;
        public string rarity;
        public List<dynamic> mainStat;
        public bool locked = false;
        public int efficiency = 0;
        public List<dynamic> subStat1;
        public List<dynamic> subStat2;
        public List<dynamic> subStat3;
        public List<dynamic> subStat4;
        public string id;
    }
}
