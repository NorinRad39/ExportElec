; =============================================================================
;  Configuration de deploiement d'ExportElec.
;
;  Seul fichier propre au projet : deploy.iss et deploy_update.ps1 viennent du
;  depot deploy-toolkit et se recopient tels quels.
;
;  Le nom n'est ecrit qu'une fois, en tete : tout ce qui en decoule est derive.
;  Les chemins relatifs partent de ce dossier.
; =============================================================================


; --- Le seul nom a changer pour un autre projet -------------------------------

#define AppName "ExportElec"


; --- Ce qui en decoule --------------------------------------------------------

; Nom technique : dossier d'installation %LOCALAPPDATA%\ExportElec, nom du setup,
; identifiant de type de fichier. Ici identique au nom affiche, faute d'espace.
#define AppSlug AppName

#define ExeName AppSlug + ".exe"

#define UpdateFolder "\\jbtec-be\meca$\topsolid\" + AppSlug

#define CertFile "cert\" + AppSlug + ".cer"
#define CertPfx "cert\" + AppSlug + ".pfx"

#define AppPublisher "Florent FABBRI"


; --- Propre a ce projet, non derivable ----------------------------------------

; Genere le 28/08/2026, et DIFFERENT de celui des autres applications. Deux
; projets qui partagent cet identifiant, et Inno prend l'un pour une mise a jour
; de l'autre : installer le second ecrase le premier sur tous les postes.
#define AppId "{{5B0E7C24-8D3A-4F16-A9C7-2E41B6D80F35}"

; ExportElec a son .csproj a la racine du depot, et non dans un sous-dossier
; comme JBT-PDFViewer ou JBTExport : le chemin de compilation remonte donc d'un
; cran de moins.
#define SourceDir "..\bin\Release"

#define IconFile "..\icone.ico"


; --- Options ------------------------------------------------------------------

; Aucune association de fichier : ExportElec n'ouvre pas de document.

; Aucun MsixPackageName : cette application n'a jamais ete empaquetee en MSIX.
; En laisser un desinstallerait le paquet d'une AUTRE application sur le poste.
 