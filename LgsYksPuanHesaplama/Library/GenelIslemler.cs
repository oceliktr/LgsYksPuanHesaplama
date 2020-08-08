using System;

namespace LgsYksPuanHesaplama.Library
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
