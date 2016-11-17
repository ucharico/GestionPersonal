using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using MiniCalc;
namespace FrmMiniCalc
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main()
        {

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            // Application.Run(new FrmMiniCalc());
            Application.Run(new MDIPrentacion());

            /*
            int nA, nB;
            string result;
            Console.WriteLine("Welcome to Mini Calc");
            Console.WriteLine("Please enter the first number to add");
            nA = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Please enter the second number to add");
            nB = Convert.ToInt32(Console.ReadLine());
            try
            {
                Calculator calc = new Calculator();
                result = (calc.Add(nA, nB)).ToString();
            }
            catch (ResultOutOfRangeException)
            {
                result = "Out of Range";
            }
            Console.WriteLine(result);
            Console.ReadLine();
            */
        }
    }
}
