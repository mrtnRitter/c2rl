## Codetwo License Reset (c2rl)

Dieses Tool automatisiert die Überwachung und das Zurücksetzen von Lizenzen im Codetwo Email Signatures 365 Dashboard.

<br>

>- Download aktuellen Release an beliebigen Speicherort
>- c2rl.exe starten
>- in config.ini Tennant ID ergänzen (aus URL emailsignatures365.codetwo.com/dashboard/tenants/**TENANT ID**/licenses)
>- c2rl.exe nochmals starten

---

Es läuft **nur unter Windows** als autostart Tray-Anwendung mit folgenden Hauptfunktionen:

- **Automatisches Lizenz-Monitoring:**  
  Überwacht kontinuierlich die aktuelle Lizenznutzung im Hintergrund.

- **Automatischer Lizenz-Reset:**  
  Setzt Lizenzzähler so bald wie möglich automatisch zurück.
  
- **Reset-Lock-Countdown:**  
  Zeigt im Tray-Menü an, wann der nächste Lizenz-Reset möglich ist.

- **Auto-Login & manuelles Login:**  
  Versucht automatisch, sich im Codetwo-Dashboard anzumelden. Falls dies fehlschlägt, wird ein manueller Login im Browser ermöglicht.

- **Statusanzeige im System-Tray:**  
  Das Tray-Icon signalisiert, wenn nur noch 2 Lizensen verfügbar sind.
  
- **Konfigurierbar:**  
  Einstellungen wie Ziel-URL, Zeitabstände und Debug-Modus werden über eine config.ini verwaltet.

- **Logging:**  
  Alle Aktionen und Fehler werden in einer rotierenden Logdatei protokolliert.

- **Einfache Bedienung:**  
  Über das Tray-Menü können Einstellungen, Infos und das Beenden der App aufgerufen werden.

---

Die Datei config.ini steuert das Verhalten des Skripts. Sie enthält folgende Einstellungen:

| Name                | Beschreibung                                                                                                |
|---------------------|------------------------------------------------------------------------------------------------------------ |
| `tenant`            | ID aus URL emailsignatures365.codetwo.com/dashboard/tenants/**TENANT ID**/licenses - manuel eintragen       |
| `user_data_dir`     | Pfad zum Chrome-Benutzerprofil-Ordner - wird automatisch gesetzt.                                           |
| `profile_dir`       | Name des Chrome-Profils bestehenden Login - wird automatisch gesetzt.                                       |
| `discover_timeout`  | Zeit in Sekunden, wie lange auf das Laden von Elementen im Browser gewartet wird.                           |
| `watchdog_timeout`  | Intervall in Sekunden, wie oft die Lizenzüberwachung ausgeführt wird.                                       |
| `debug`             | `True` für Fenstermodus, `False` für fensterlosen Hintergrundmodus                                          |

> **Hinweis:**  
> config.ini wird beim ersten Start automatisch erstellt und kann nachträglich angepasst werden.  
> Pfad- und Profilangaben werden automatisch ermittelt.

---

Für die Ausführung als Skript werden der **Chrome Browser** und diese Module benötigt: 
- **selenium**
- **pystray**
- **pillow (PIL)**
- **pywin32**

Installation:
```
pip install selenium pystray pillow win32com.client
```

