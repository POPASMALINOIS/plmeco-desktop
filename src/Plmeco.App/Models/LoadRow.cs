using System;

namespace Plmeco.App.Models;

public class LoadRow
{
    public int Id { get; set; }
    public string Transportista { get; set; } = "";
    public string Matricula { get; set; } = "";
    public string Destino { get; set; } = "";
    public string Muelle { get; set; } = "";
    public string Estado { get; set; } = "";
    public TimeSpan? LlegadaReal { get; set; }
    public TimeSpan? SalidaReal { get; set; }
    public TimeSpan? SalidaTope { get; set; }
    public bool Lex { get; set; }
    public string Incidencias =>
        (SalidaReal.HasValue && SalidaTope.HasValue && SalidaReal > SalidaTope)
        ? "RETRASO TRANSPORTISTA" : string.Empty;
}
