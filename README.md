# Citrix Printeroverzicht

PowerShell-tool met een grafische interface (WPF) die alle printers in de huidige Citrix-sessie toont, opgemaakt in de huisstijl van het Streekziekenhuis Koningin Beatrix (SKB).

## Versie

Huidige versie: **26112301**. Het versienummer heeft het formaat `JJMMDDnn` (datum + volgnummer van die dag). Het staat bovenin het script (`$ScriptVersion`, met versiegeschiedenis) en in de statusbalk van de applicatie.

## Starten

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File .\Show-CitrixPrinters.ps1
```

- Geen beheerdersrechten nodig.
- Werkt met Windows PowerShell 5.1 en PowerShell 7.
- Eén los bestand: het logo zit in het script ingebouwd.

## Functies

- **Sessie-informatie:** gebruiker en server, clientnaam en IP-adres van het eindapparaat, en de huidige standaardprinter. Het script herkent automatisch of het in een Citrix-sessie draait.
- **Printerlijst:** standaardprinter (ster), naam, type (Citrix clientprinter, netwerkprinter, lokaal, virtueel), status, bron, locatie en opmerking (het veld Comment van de printer).
- **Zoeken en filteren** op naam, bron, locatie of opmerking, en op type.
- **Acties:** vernieuwen (ook met F5), als standaardprinter instellen, afdrukwachtrij openen (ook door te dubbelklikken), printernaam kopiëren (via het rechtsklikmenu).
- **Automatisch vernieuwen** elke 30 seconden. Handig omdat Citrix clientprinters soms pas na het inloggen verschijnen.

## Hoe het werkt

| Onderdeel | Bron |
|---|---|
| Printers | `Get-CimInstance -ClassName Win32_Printer` |
| Standaard instellen | `Invoke-CimMethod -MethodName SetDefaultPrinter` |
| Citrix-sessie, client en IP | `HKLM:\SOFTWARE\Citrix\Ica\Session\<sessie-id>\Connection`, `$env:SESSIONNAME`, `$env:CLIENTNAME` |

Een printer telt als **Citrix clientprinter** bij een `CLIENT:`-poort, `(from …)` of `(van …)` in de naam, of de Citrix Universal Printer-driver. Een printer als `\\server\printer` telt als **netwerkprinter**.
