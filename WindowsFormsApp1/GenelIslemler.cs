using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
   public static class GenelIslemler
    {
        public static int ToInt32(this object sayi)
        {
            try
            {
                if (sayi == null) throw new Exception();
                int x = Convert.ToInt32(sayi);
                return x;
            }
            catch (Exception)
            {
                return 0;
            }
        }
        public static double ToDouble(this object sayi)
        {
            try
            {
                if (sayi == null) throw new Exception();
                double x = Convert.ToDouble(sayi);
                return x;
            }
            catch (Exception)
            {
                return 0;
            }
        }
        public static decimal ToDecimal(this object sayi)
        {
            try
            {
                if (sayi == null) throw new Exception();
                decimal x = Convert.ToDecimal(sayi);
                return x;
            }
            catch (Exception)
            {
                return 0;
            }
        }
    }
}
