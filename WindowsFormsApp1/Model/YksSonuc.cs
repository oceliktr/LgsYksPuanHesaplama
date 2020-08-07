using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.Model
{
   public class YksSonuc
    {
        public string Kategori { get; set; }
        public string Tckimlik { get; set; }
        public string IlceAdi { get; set; }
        public string OkulAdi { get; set; }
        public int OgrenciSayisi { get; set; }
        public string AdiSoyadi { get; set; }

        public int TurkDiliDogru { get; set; }
        public int TurkDiliYanlis { get; set; }
        public decimal TurkDiliNet { get; set; }
        public bool TurkDiliGirdi { get; set; }
        public int Tarih1Dogru { get; set; }
        public int Tarih1Yanlis { get; set; }
        public decimal Tarih1Net { get; set; }
        public bool Tarih1Girdi { get; set; }
        public int Cog1Dogru { get; set; }
        public int Cog1Yanlis { get; set; }
        public decimal Cog1Net { get; set; }
        public bool Cog1Girdi { get; set; }
        public int Tarih2Dogru { get; set; }
        public int Tarih2Yanlis { get; set; }
        public decimal Tarih2Net { get; set; }
        public bool Tarih2Girdi { get; set; }
        public int Cog2Dogru { get; set; }
        public int Cog2Yanlis { get; set; }
        public decimal Cog2Net { get; set; }
        public bool Cog2Girdi { get; set; }
        public int FelsefeDogru { get; set; }
        public int FelsefeYanlis { get; set; }
        public decimal FelsefeNet { get; set; }
        public bool FelsefeGirdi { get; set; }
        public int DinDogru { get; set; }
        public int DinYanlis { get; set; }
        public decimal DinNet { get; set; }
        public bool DinGirdi { get; set; }
        public int MatDogruAyt { get; set; }
        public int MatYanlisAyt { get; set; }
        public decimal MatNetAyt { get; set; }
        public bool MatAytGirdi { get; set; }
        public int FizikDogru { get; set; }
        public int FizikYanlis { get; set; }
        public decimal FizikNet { get; set; }
        public bool FizikGirdi { get; set; }
        public int KimyaDogru { get; set; }
        public int KimyaYanlis { get; set; }
        public decimal KimyaNet { get; set; }
        public bool KimyaGirdi { get; set; }
        public int BiyolojiDogru { get; set; }
        public int BiyolojiYanlis { get; set; }
        public decimal BiyolojiNet { get; set; }
        public bool BiyolojiGirdi { get; set; }
        public decimal ToplamAytNet { get; set; }
        public int TurkceDogru { get; set; }
        public int TurkceYanlis { get; set; }
        public decimal TurkceNet { get; set; }
        public bool TurkceGirdi { get; set; }
        public int SosyalBDogru { get; set; }
        public int SosyalBYanlis { get; set; }
        public decimal SosyalBNet { get; set; }
        public bool SosyalBGirdi { get; set; }
        public int FenDogru { get; set; }
        public int FenYanlis { get; set; }
        public decimal FenNet { get; set; }
        public bool FenGirdi { get; set; }

        public int MatDogruTyt { get; set; }
        public int MatYanlisTyt { get; set; }
        public decimal MatNetTyt { get; set; }
        public bool MatTytGirdi { get; set; }
        public decimal TYTPuanYerl { get; set; }
        public decimal SayisalPuanYerl { get; set; }
        public decimal SozelPuanYerl { get; set; }
        public decimal EsitAgirlikPuanYerl { get; set; }
        public decimal YabanciDilPuanYerl { get; set; }
        public decimal TYTPuanYuzde { get; set; }
        public decimal SayisalPuanYuzde { get; set; }
        public decimal SozelPuanYuzde { get; set; }
        public decimal EsitAgirlikPuanYuzde { get; set; }
        public decimal YabanciDilPuanYuzde { get; set; }
        public decimal ToplamTytNet { get; set; }
        
    }
}
