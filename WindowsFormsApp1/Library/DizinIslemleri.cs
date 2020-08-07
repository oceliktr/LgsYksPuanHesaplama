using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WindowsFormsApp1.Model;

namespace WindowsFormsApp1.Library
{
   public class DizinIslemleri
    {
        public List<DosyaInfo> DizindekiDosyalariListele(string dizinAdresi)
        {
            DirectoryInfo dizin = new DirectoryInfo(dizinAdresi);
            var dosyalar = dizin.GetFiles("*.*", SearchOption.AllDirectories); //Alt dizindekileri de listelemek için SearchOption.AllDirectories kullan
            List<DosyaInfo> list = new List<DosyaInfo>();
            foreach (FileInfo dsy in dosyalar)
            {
                DosyaInfo lst = new DosyaInfo(dsy.Name, dizinAdresi, dsy.CreationTime, dsy.DirectoryName);
                list.Add(lst);
            }
            return list;
        }
    }
}
