using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SyncTestApp
{
    public class Program
    {
        static void Main(string[] args)
        {
            if (args.Count() >= 4)
            {
                string siteType = args[0];
                string id = args[1];
                string code = args[2];
                string name = args[3];

                SyncService.SyncOpleidingscatalogusServiceClient client = new SyncService.SyncOpleidingscatalogusServiceClient();
                SyncService.ActieRapport rapport = null;
                if (args[0].ToLower().Equals("edu"))
                {
                    SyncService.OpleidingVal opleiding = new SyncService.OpleidingVal();
                    opleiding.Id = id;
                    opleiding.Code = code;
                    opleiding.Naam = name;

                    if (args.Count() == 6)
                    {
                        string eduWorkspace = args[4];
                        string eduType = args[5];
                        opleiding.EduWorkSpace = eduWorkspace;
                        opleiding.EduType = eduType;
                    }
                    rapport = client.DoeOnbepaaldeActieOpleiding(opleiding);
                } // end if edu
                if (args[0].ToLower().Equals("mod"))
                {
                    SyncService.ModuleVal module = new SyncService.ModuleVal();
                    module.Id = id;
                    module.Code = code;
                    module.Naam = name;

                    if (args.Count() == 6)
                    {
                        string eduCode = args[4];
                        string linkedModule = args[5];
                        module.EduCode = eduCode;
                        module.LinkedModule = linkedModule;
                    }
                    rapport = client.DoeOnbepaaldeActieModule(module);
                    foreach (string trace in rapport.Berichten)
                    {
                        Console.WriteLine(trace);
                    }

                } // end if mod
                foreach (string trace in rapport.Berichten)
                {
                    Console.WriteLine(trace);
                }
            } // end if
            else
            {
                Console.WriteLine("Syntax: <edu or mod> <id> <code> <name> (<eduWorkspace>|<eduCode>) (<eduType>|<linkedModule>)");
            } // end else
        }
    } // end c
} // end ns
