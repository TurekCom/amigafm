# Changelog

## Unreleased

## 0.1.2 - 2026-04-30

### Dodano

- Niewidoczny podgląd multimediów pod `F3`, uruchamiany i zatrzymywany bez otwierania zewnętrznego okna.
- Dynamiczne ładowanie `bass.dll` i pluginów `bass*.dll` dla podglądu audio z plików lokalnych oraz z zasobów sieciowych.
- Strumieniowanie podglądu multimediów z zasobów HTTP, HTTPS, FTP, FTPS i WebDAV, gdy zasób nie wymaga hasła.
- Bezokienny fallback przez `ffplay` lub `mpv` dla plików video, których BASS nie potrafi otworzyć.
- Zasoby sieciowe HTTP i HTTPS tylko do odczytu z parsowaniem typowych listingów katalogów.
- Edycję uprawnień SFTP przez `chmod` i `chown`, także rekurencyjnie oraz z ponowieniem przez sudo/root przy odmowie dostępu.

### Zmieniono

- Dialog uprawnień SFTP ma pola wyboru odczyt, zapis i wykonanie dla właściciela, grupy oraz innych.
- Dialog dodawania i edycji połączenia pokazuje pole klucza SSH tylko dla SFTP.
- Parser listingów HTTP/HTTPS odczytuje rozmiary także z tabel HTML, w tym wartości z przecinkiem dziesiętnym.
- Komunikaty błędów HTTP 403 są bardziej jednoznaczne, gdy serwer pokazuje plik w listingu, ale odmawia pobrania.

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
