using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GPACalc
{
    class Help
    {
        /// <summary>
        /// Mail to Clive
        /// </summary>
        public static void ContactClive()
        {
            System.Diagnostics.Process.Start("mailto:Clive.DM@outlook.com?subject=FromGPACalc");
        }
    }
}
