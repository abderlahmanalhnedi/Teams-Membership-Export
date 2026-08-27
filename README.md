# Microsoft Teams Membership Export

A small PowerShell utility for Microsoft 365 / Teams administrators that exports all Microsoft Teams memberships of a selected user, including the user's role in each Team, to an Excel workbook.

**Version:** 1.0.0  
**Author:** Abdelrahman Al-Hnedi  
**Last updated:** 2026-08-27

---

## Deutsch

### Funktionen

Das Skript führt den Administrator interaktiv durch den gesamten Ablauf:

1. Admin-E-Mail / UPN eingeben.
2. Admin-Passwort sicher über `Get-Credential` eingeben.
3. E-Mail / UPN des Zielbenutzers eingeben.
4. Verbindung mit Microsoft Teams herstellen.
5. Alle Teams-Mitgliedschaften des Zielbenutzers auslesen.
6. Rolle je Team ermitteln:
   - `Owner`
   - `Member`
   - `Guest`
7. Bericht als echte `.xlsx`-Datei auf dem Desktop erstellen.
8. Excel-Datei automatisch öffnen.

Die erzeugte Arbeitsmappe enthält:

- **Teams** – Teamname, Rolle und GroupId.
- **Info** – Zielbenutzer, verwendetes Admin-Konto, Statistik, Exportzeitpunkt, PowerShell-Version und MicrosoftTeams-Modulversion.

### Voraussetzungen

- Windows
- Windows PowerShell **5.1** oder PowerShell **7.2+**
- Zugriff auf Microsoft Teams / Microsoft 365 mit einem Konto, das die benötigten Berechtigungen besitzt
- Internetzugriff auf Microsoft 365 und optional PowerShell Gallery

Benötigte PowerShell-Module:

- `MicrosoftTeams`
- `ImportExcel`

Fehlende Module können vom Skript für den aktuellen Benutzer automatisch installiert werden.

### Installation

Repository herunterladen oder klonen und anschließend das Skript starten:

```powershell
cd C:\Pfad\zum\Repository
.\Teams-Membership-Export.ps1
```

Eine Installation von Microsoft Excel ist für die Erstellung der XLSX-Datei nicht erforderlich. Das Skript verwendet das PowerShell-Modul `ImportExcel`.

### Anmeldung und MFA

Das Skript fragt das Admin-Passwort über `Get-Credential` ab. Das Passwort wird **nicht** im Skript gespeichert und **nicht** in den Excel-Bericht geschrieben.

Die Anmeldeversuche erfolgen in dieser Reihenfolge:

1. Benutzername + Passwort mit `-DisableWAM`, wenn die installierte MicrosoftTeams-Version den Parameter unterstützt.
2. Benutzername + Passwort mit der Standardauthentifizierung des MicrosoftTeams-Moduls.
3. Interaktive Microsoft-Anmeldung.
4. Gerätecode-Anmeldung als zusätzliche Alternative, wenn unterstützt.

Bei Konten mit **MFA** oder **Conditional Access** kann Microsoft eine interaktive Anmeldung verlangen. In diesem Fall kann eine reine Benutzername-/Passwort-Anmeldung nicht ausreichen.

### Beispiel

Nach dem Start wird beispielsweise nach folgenden Werten gefragt:

```text
Admin-Konto:
admin@contoso.com

Zielbenutzer:
user@contoso.com
```

Das Ergebnis wird beispielsweise unter folgendem Namen auf dem Desktop gespeichert:

```text
John Doe_Teams_2026-08-27_091500.xlsx
```

### Beispielausgabe

| TeamName | Role | GroupId |
|---|---|---|
| IT Operations | Owner | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` |
| Microsoft 365 | Member | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` |
| Company News | Member | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` |

### Sicherheit

Das Repository enthält keine organisationsspezifischen Konten, Passwörter, Tenant-IDs oder Secrets.

Das Admin-Passwort wird während der Ausführung als `SecureString` innerhalb eines `PSCredential`-Objekts behandelt und nicht dauerhaft gespeichert.

Vor einem Commit sollte trotzdem immer geprüft werden:

```powershell
git diff
git status
```

Damit lässt sich kontrollieren, dass keine internen oder vertraulichen Daten versehentlich hinzugefügt wurden.

### Bekannte Hinweise

- PowerShell ISE wird nicht empfohlen. Verwende bevorzugt Windows PowerShell Console, PowerShell 7 oder Visual Studio Code mit der PowerShell-Erweiterung.
- Die verfügbaren Teams-Daten hängen von den Berechtigungen des angemeldeten Administrators ab.
- Microsoft kann den Authentifizierungsablauf durch MFA, Conditional Access oder Änderungen im MicrosoftTeams-Modul beeinflussen.
- `-DisableWAM` ist ein temporärer Kompatibilitätsparameter des MicrosoftTeams-Moduls und kann in einer späteren Version wieder entfernt werden.

---

## English

### Features

The script provides an interactive workflow:

1. Enter the administrator UPN/e-mail address.
2. Enter the administrator password securely with `Get-Credential`.
3. Enter the target user's UPN/e-mail address.
4. Connect to Microsoft Teams.
5. Retrieve all Teams memberships of the target user.
6. Determine the role in each Team:
   - `Owner`
   - `Member`
   - `Guest`
7. Create a real `.xlsx` report on the current user's Desktop.
8. Open the generated workbook automatically.

The workbook contains two worksheets:

- **Teams** – Team name, role, and GroupId.
- **Info** – Target user, administrator account, statistics, export timestamp, PowerShell version, and MicrosoftTeams module version.

### Requirements

- Windows
- Windows PowerShell **5.1** or PowerShell **7.2+**
- A Microsoft 365 / Teams account with sufficient permissions
- Internet access to Microsoft 365 and optionally PowerShell Gallery

Required modules:

- `MicrosoftTeams`
- `ImportExcel`

Missing modules can be installed automatically for the current user.

### Usage

```powershell
cd C:\Path\To\Repository
.\Teams-Membership-Export.ps1
```

Microsoft Excel itself is not required to create the workbook.

### Authentication

The administrator password is requested through `Get-Credential`. It is not hard-coded, written to disk, or included in the exported workbook.

If password-based authentication cannot be used because of MFA, Conditional Access, or the current MicrosoftTeams authentication flow, the script can fall back to interactive or device-code authentication.

### Security

No tenant-specific usernames, passwords, tenant IDs, or secrets are stored in the repository.

Before publishing changes, review them with:

```powershell
git diff
git status
```

### Notes

- PowerShell ISE is not recommended. Prefer Windows PowerShell Console, PowerShell 7, or Visual Studio Code with the PowerShell extension.
- Results depend on the permissions of the signed-in administrator.
- Authentication behavior may change with Microsoft Teams PowerShell updates and tenant security policies.

---

## Microsoft documentation

- [Install Microsoft Teams PowerShell](https://learn.microsoft.com/microsoftteams/teams-powershell-install)
- [Connect-MicrosoftTeams](https://learn.microsoft.com/powershell/module/teams/connect-microsoftteams)
- [Get-Team](https://learn.microsoft.com/powershell/module/teams/get-team)
- [Get-TeamUser](https://learn.microsoft.com/powershell/module/teams/get-teamuser)

---

## Repository structure

```text
Microsoft-Teams-Membership-Export/
├── Teams-Membership-Export.ps1
└── README.md
```
