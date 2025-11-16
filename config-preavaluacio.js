// Configuració de la preavaluació
// Pots adaptar aquests valors a altres cursos / fulls / taules sense tocar l'index.html

window.PREAVAL_CONFIG = {
  // Azure AD / Microsoft Entra
  clientId: "1892e1b0-6724-4cfe-b0aa-83d746e755e8",
  tenantId: "f6fbd66a-02b9-4804-8f74-b16600cf5e00",

  // Si en el futur vols utilitzar un site de SharePoint en lloc de OneDrive personal,
  // posa useSharePointSite: true i omple hostname + sitePath
  useSharePointSite: false,
  sharePoint: {
    hostname: "https://gmqualitytechnologysl365.sharepoint.com",
    sitePath: "/sites/DIGITECHBarcelona/SitePages/"
  },

  // Fitxer de referències (Alumnes, Assignacions, Professors)
  refsFilePath: "/tutoria/Preavaluacio/referencies.xlsx",
  refsTables: {
    alumnes: "Alumnes",
    assign: "Assignacions",
    professors: "Professors"
  },

  // Fitxer amb les preavaluacions guardades
  preFilePath: "/tutoria/Preavaluacio/preavaluacio_dades.xlsx",
  preTableName: "Respostes",

  // Opcional: si la deixes buida, l'index.html calcularà automàticament el redirectUri
  redirectUri: ""
};
