using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SeleniumSample
{
   public class Commons
    {
        public static void Wait(int second)
        {
            for (int i = 0; i < second; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    System.Threading.Thread.Sleep(100);
                    Application.DoEvents();
                }
            }
        }
        public static void Wait()
        {
            System.Threading.Thread.Sleep(100);
            Application.DoEvents();
            System.Threading.Thread.Sleep(100);
            Application.DoEvents();
        }
    }
}
