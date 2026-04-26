# Changelog

## Unreleased

### Dodano

- Testową obsługę Dropbox i Google Drive jako zasobów sieciowych.
- Logowanie OAuth dla Dropbox i Google Drive z otwarciem przeglądarki systemowej, lokalnym callbackiem `127.0.0.1` i PKCE.
- Ukrywanie pól serwerowych w dialogu połączenia po wybraniu Dropbox albo Google Drive.
- Dropbox przez API v2: listowanie, pobieranie, wysyłanie, tworzenie folderów, usuwanie i przenoszenie.
- Google Drive przez Drive API v3: ścieżkowy dostęp do plików i folderów, pobieranie, wysyłanie, tworzenie folderów, usuwanie i przenoszenie.

### Znane ograniczenia

- Dropbox i Google Drive wymagają wbudowania identyfikatora aplikacji OAuth przez wydawcę builda.
- Dropbox wymaga dodania redirect URI `http://127.0.0.1:53682/callback` w konsoli aplikacji.
- Pliki Google Workspace wymagające eksportu nie są jeszcze pobierane.
- Nazwy z ukośnikiem w Google Drive nie są jeszcze mapowane na bezpieczne nazwy panelu.

## 0.1.1 - 2026-04-24

### Zmieniono

- Panele przy pierwszym uruchomieniu pokazują listę dysków zamiast katalogu roboczego programu.
- Program zapisuje ostatnią lokalizację lewego i prawego panelu.
- Jeśli zapisana lokalizacja panelu jest niedostępna przy następnym uruchomieniu, panel wraca do listy dysków.
- Instalator Inno Setup instaluje program w `Program Files` i wymaga uprawnień administratora.

## 0.1.0 - 2026-04-24

Pierwsze wydanie prototypowe Amiga FM.

### Dodano

- Natywny interfejs Win32 inspirowany Amiga Workbench.
- Dwupanelowy menedżer plików z obsługą klawiatury.
- Komunikaty dla NVDA przez NVDA Controller Client.
- Operacje kopiowania, przenoszenia, usuwania, zmiany nazwy i tworzenia katalogów.
- Obsługę schowka Windows dla plików i katalogów.
- Zasoby sieciowe SFTP, SMB, FTP, FTPS, WebDAV i NFS.
- Skanowanie usług w sieci lokalnej i cache wyników.
- Obsługę chronionych lokalizacji SFTP przez sudo/su.
- Wyszukiwanie z wyrażeniami regularnymi.
- Ulubione katalogi i ulubione pliki.
- Obsługę archiwów i obrazów przez 7-Zip.
- Tworzenie archiwów `7z`, `zip`, `tar`, `tar.gz`, `tar.bz2`, `tar.xz`, `gzip`, `bzip2`, `xz`, `wim`.
- Tworzenie i sprawdzanie sum kontrolnych SHA-256.
- Instalator Inno Setup ze skrótem na pulpicie.

### Znane ograniczenia

- Tworzenie RAR nie jest dostępne, ponieważ 7-Zip obsługuje RAR tylko do odczytu.
- Zaawansowane operacje wewnątrz niektórych obrazów dysków zależą od możliwości zainstalowanego 7-Zip.
- 7-Zip musi być zainstalowany osobno, jeśli użytkownik chce korzystać z archiwów.
