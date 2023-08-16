namespace App.Controllers
{
    internal class ResponseViewModel
    {
        public int Premio { get; set; }
        public int Quantidade { get; set; }
        public ParticipanteViewModel Ganhador1 { get; set; }
        public ParticipanteViewModel Ganhador2 { get; set; }
        public ParticipanteViewModel Ganhador3 { get; set; }
    }
}