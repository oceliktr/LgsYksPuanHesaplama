namespace LgsYksPuanHesaplama.Model
{
    public class LgsSonuc
    {
        public string Tckimlik { get; set; }
        public string IlceAdi { get; set; }
        public string OkulAdi { get; set; }
        public int OgrenciSayisi { get; set; }
        public string AdiSoyadi { get; set; }
        public decimal SinavPuani { get; set; }
        public decimal YuzdelikDilim { get; set; }
        public decimal ToplamNet { get; set; }
        public int TurkceDogru { get; set; }
        public int TurkceYanlis { get; set; }
        public decimal TurkceNet { get; set; }
        public int MatDogru { get; set; }
        public int MatYanlis { get; set; }
        public decimal MatNet { get; set; }
        public int FenDogru { get; set; }
        public int FenYanlis { get; set; }
        public decimal FenNet { get; set; }
        public int InkDogru { get; set; }
        public int InkYanlis { get; set; }
        public decimal InkNet { get; set; }
        public int DinDogru { get; set; }
        public int DinYanlis { get; set; }
        public decimal DinNet { get; set; }
        public int IngDogru { get; set; }
        public int IngYanlis { get; set; }
        public decimal IngNet { get; set; }
        public string Aciklama { get; set; }
    }
}
