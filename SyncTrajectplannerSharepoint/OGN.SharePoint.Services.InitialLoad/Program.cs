using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OGN.Sharepoint.Services;
using System.IO;
using System.Configuration;

namespace OGN.SharePoint.Services.InitialLoad
{
    class Program
    {
        private static bool DoEdu = false;
        private static bool DoMod = false;
        private static bool DoRelations = false;

        static void Main(string[] args)
        {
            SyncEduSitesService svc = new SyncEduSitesService(false);
            //svc.FixSiteNames();
            EduProgramme edu = new EduProgramme();
            edu.Code = "1572.100";
            edu.Name = "Test Name5";
            edu.Id = "vio1";
            edu.EduType = "TEST";

            EduProgrammeVal eduVal = new EduProgrammeVal();
            eduVal.Code = edu.Code;
            eduVal.Id = edu.Id;
            eduVal.Name = edu.Name;
            eduVal.EduType = edu.EduType;

            //svc.DoUndeterminedAction(eduVal);

            Module mod = new Module();
            mod.Code = "99999";
            mod.Id = "op1";
            mod.Name = "Test Module LOI4";
            mod.EduCode = "1572.100";
            

            ModuleVal modVal = new ModuleVal();
            modVal.Code = mod.Code;
            modVal.Id = mod.Id;
            modVal.Name = mod.Name;
            modVal.EduCode = mod.EduCode;
            modVal.LinkedModule = mod.LinkedModule;

            svc.DoUndeterminedAction(modVal);


            if (args.Length == 0)
            {
                DoEdu = true;
                DoMod = true;
                DoRelations = true;
                Program.Create();
            }
            else
            {
                if (args[0].Equals("ALLESWEG")) { Program.AllesWeg(args[1]); }

                if (args[0].ToLower().Equals("edu"))
                {
                    DoEdu = true;
                    Program.Create();
                }

                if (args[0].ToLower().Equals("mod"))
                {
                    DoMod = true;
                    Program.Create();
                }

                if (args[0].ToLower().Equals("rel"))
                {
                    DoRelations = true;
                    Program.Create();
                }
            }
        }

        static void AllesWeg(string s)
        {
            SyncEduSitesService svc = new SyncEduSitesService();
            svc.DeleteSubsites(s);
        }
        static void Create()
        {
            SyncEduSitesService svc = new SyncEduSitesService(false);
            char[] delim = { ';' };

            File.WriteAllText("log.csv", "");
            using (FileStream logfile = File.OpenWrite("log.csv"))
            {
                using (StreamWriter log = new StreamWriter(logfile))
                {
                    if (DoEdu)
                    {
                        Console.WriteLine("Opleidingen...");
                        string[] opl = File.ReadAllLines("Opleidingen.csv", Encoding.Default);
                        for (int i = 1; i < opl.Length; i++)
                        {
                            log.WriteLine("Opleiding;" + opl[i]);
                            string[] val = opl[i].Split(delim, 5);
                            EduProgrammeVal edu = new EduProgrammeVal();
                            edu.Id = val[0];
                            edu.Code = val[1];
                            edu.Name = val[2];
                            edu.EduWorkSpace = val[3];
                            edu.EduType = val[4];
                            try
                            {
                                Console.WriteLine("{3} out of {4}: Creating opleidingssite: {0}, type: {1}, code: {2}", edu.Name, edu.EduType, edu.Id, i.ToString(), opl.Length.ToString());
                                DateTime startTime = DateTime.Now;
                                svc.Create(edu);
                                DateTime endTime = DateTime.Now;
                                TimeSpan diff = endTime.Subtract(startTime);
                                Console.WriteLine("Opleiding Site creation took: {0} seconds", diff.TotalSeconds.ToString());
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("Error: {0}", e.Message);
                                log.WriteLine(edu.Id + " (opleiding);" + e.Message);
                            }
                        }
                    }

                    if (DoMod)
                    {
                        Console.WriteLine("Modules...");
                        string[] mods = File.ReadAllLines("Modules.csv", Encoding.Default);
                        for (int i = 1; i < mods.Length; i++)
                        {
                            log.WriteLine("Module;" + mods[i]);
                            string[] val = mods[i].Split(delim, 5);
                            ModuleVal mod = new ModuleVal();
                            mod.Id = val[0];
                            mod.Code = val[1];
                            mod.Name = val[2];
                            mod.EduCode = val[3];
                            mod.LinkedModule = val[4];
                            try
                            {
                                DateTime startTime = DateTime.Now;
                                Console.WriteLine("{3} out of {4}: Creating modulesite: {0}, code: {1}", mod.Name,  mod.Id, i.ToString(), mods.Length.ToString());
                                svc.Create(mod);
                                DateTime endTime = DateTime.Now;
                                TimeSpan diff = endTime.Subtract(startTime);
                                Console.WriteLine("Module Site creation took: {0} seconds", diff.TotalSeconds.ToString());

                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("Error: {0}", e.Message);
                                log.WriteLine(mod.Id + " (module);" + e.Message);
                            }
                        }
                    }

                    if (DoRelations)
                    {
                        Console.WriteLine("Relaties...");
                        string[] rels = File.ReadAllLines("Relaties.csv", Encoding.Default);
                        for (int i = 1; i < rels.Length; i++)
                        {
                            log.WriteLine("Relatie;" + rels[i]);
                            string[] val = rels[i].Split(delim, 2);
                            Link rel = new Link();
                            rel.EduProgramme = new EduProgrammeRef(val[0]);
                            rel.Module = new ModuleRef(val[1]);
                            try
                            {
                                Console.WriteLine("{2} out of {3}: Creating relation Opleiding: {0}, Module: {1}", rel.EduProgramme.Id, rel.Module.Id, i.ToString(), rels.Length.ToString());
                                DateTime startTime = DateTime.Now;
                                svc.Create(rel);
                                DateTime endTime = DateTime.Now;
                                TimeSpan diff = endTime.Subtract(startTime);
                                Console.WriteLine("Relation creation took: {0} seconds", diff.TotalSeconds.ToString());
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("Error: {0}", e.Message);
                                log.WriteLine(rel.EduProgramme.Id + " (opleiding): " + rel.Module.Id + " (module);" + e.Message);
                            }
                        }
                    }
                } // end log
            } // end filestream
        } // end main
    } // end class
} // end ns
