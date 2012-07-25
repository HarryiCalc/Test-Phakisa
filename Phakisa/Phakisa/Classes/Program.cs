using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Phakisa
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        public static string[] Args = null;

        [STAThread]
        static void Main(string[] args)
        {
            //Program.Args = args;
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);

            ////string[] test = Program.Args;
            //string[] test = new string[1] { "201002-FREEGOLD-JJ-USERA-STOPE-TEAM-Development" };

            ////split parameters  
            //string Period = test[0].ToString().Substring(0, test[0].ToString().Trim().IndexOf("-"));
            //test[0] = test[0].Trim().Substring(test[0].ToString().Trim().IndexOf("-") + 1);
            //string Region = test[0].ToString().Substring(0, test[0].ToString().Trim().IndexOf("-"));
            //test[0] = test[0].Trim().Substring(test[0].ToString().Trim().IndexOf("-") + 1);
            //string BussUnit = test[0].ToString().Substring(0, test[0].ToString().Trim().IndexOf("-"));
            //test[0] = test[0].Trim().Substring(test[0].ToString().Trim().IndexOf("-") + 1);
            //string Userid = test[0].ToString().Substring(0, test[0].ToString().Trim().IndexOf("-"));
            //test[0] = test[0].Trim().Substring(test[0].ToString().Trim().IndexOf("-") + 1);
            //string MiningType = test[0].ToString().Substring(0, test[0].ToString().Trim().IndexOf("-"));
            //test[0] = test[0].Trim().Substring(test[0].ToString().Trim().IndexOf("-") + 1);
            //string BonusType = test[0].ToString().Substring(0, test[0].ToString().Trim().IndexOf("-"));
            //test[0] = test[0].Trim().Substring(test[0].ToString().Trim().IndexOf("-") + 1);
            //string Environment = test[0].ToString().Substring(0).Trim();

            //switch (MiningType.Trim() + " " + BonusType.Trim())
            //{
            //    case "STOPE TEAM":
            //        scrTeamD TeamS = new scrTeamD();
            //        TeamS.scrTeamDLoad(Period, Region, BussUnit, Userid, MiningType, BonusType, Environment);
            //        TeamS.ShowDialog();
            //        //Application.Run(new scrTeamD.scrTeamDLoad(Period, Region, BussUnit, Userid, MiningType, BonusType, Environment));
            //        break;
                
            //}

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new scrLogon());

        }
    }
}
