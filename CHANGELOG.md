# Changelog

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
