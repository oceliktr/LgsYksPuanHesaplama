using System;

namespace LgsYksPuanHesaplama.Model
{
    public class DosyaInfo
    {
        public string DosyaAdi { get; set; }
        public string DosyaYolu { get; set; }
        public string DizinAdresi { get; set; }
        public DateTime DosyaOlusturmaTarihi { get; set; }
        public DosyaInfo()
        {
        }
        public DosyaInfo(string dosyaAdi, string dosyaYolu, DateTime dosyaOlusturmaTarihi, string dizinAdresi)
        {
            DosyaAdi = dosyaAdi;
            DosyaYolu = dosyaYolu;
            DosyaOlusturmaTarihi = dosyaOlusturmaTarihi;
            DizinAdresi = dizinAdresi;
        }
    }
}
