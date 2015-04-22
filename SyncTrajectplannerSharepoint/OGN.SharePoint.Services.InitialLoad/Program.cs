﻿using System;
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
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Program.Create();
            }
            else
            {
                if (args[0].Equals("ALLESWEG")) { Program.AllesWeg(args[1]); }
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
                        edu.LOISite = val[3];
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
                            log.WriteLine(edu.Id + " (opleiding);" + e.Message);
                        }
                    }
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
                        mod.LOISite = val[3];
                        mod.EduType = val[4];
                        try
                        {
                            DateTime startTime = DateTime.Now;
                            Console.WriteLine("{2} out of {3}: Creating modulesite: {0}, type: {1}, code: {2}", mod.Name, mod.EduType, mod.Id, i.ToString(), mods.Length.ToString());
                            svc.Create(mod);
                            DateTime endTime = DateTime.Now;
                            TimeSpan diff = endTime.Subtract(startTime);
                            Console.WriteLine("Module Site creation took: {0} seconds", diff.TotalSeconds.ToString());

                        }
                        catch (Exception e)
                        {
                            log.WriteLine(mod.Id + " (module);" + e.Message);
                        }
                    }
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
                            log.WriteLine(rel.EduProgramme.Id + " (opleiding): " + rel.Module.Id + " (module);" + e.Message);
                        }
                    }
                }
            }
        }
    }
}
