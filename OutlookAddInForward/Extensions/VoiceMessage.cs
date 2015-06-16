namespace OutlookAddInForward.Extensions
{
    using System;

    public class VoiceMessage
    {
        public string Anexo { get; set; }
        public string UsuarioAnexo { get; set; }
        public bool LlamadaInterna { get; set; }
        public string NumeroTelefono { get; set; }
        public DateTime FechaLlamada { get; set; }
    }
}
