# Amiga FM

Amiga FM to lekki, natywny menedżer plików dla Windows napisany w Rust. Interfejs jest inspirowany klasycznym Amiga Workbench: czarne tło, żółte etykiety i zaznaczenie w negatywie. Program jest projektowany z myślą o obsłudze przez NVDA.

## Najważniejsze funkcje

- Dwa panele plików z obsługą klawiatury.
- Dostępność przez natywne kontrolki Win32 i komunikaty NVDA Controller Client.
- Operacje na plikach: kopiowanie, przenoszenie, usuwanie, zmiana nazwy, tworzenie katalogów.
- Schowek Windows: `Ctrl+C`, `Ctrl+X`, `Ctrl+V` między Amiga FM i innymi aplikacjami.
- Zasoby sieciowe: SFTP, SMB, FTP, FTPS, WebDAV i NFS.
- Wyszukiwanie lokalne i rekurencyjne z wyrażeniami regularnymi.
- Ulubione katalogi i pliki.
- Obsługa archiwów i obrazów przez 7-Zip: otwieranie, wypakowywanie, dodawanie plików, usuwanie z archiwum.
- Tworzenie archiwów: `7z`, `zip`, `tar`, `tar.gz`, `tar.bz2`, `tar.xz`, `gzip`, `bzip2`, `xz`, `wim`.
- Odczyt wielu formatów, między innymi `rar`, `zipx`, `jar`, `apk`, `cab`, `iso`, `vhd`, `vhdx`, `vdi`, `vmdk`, `qcow2`, `dmg`, `squashfs`, `wim`, `swm`, `deb`, `rpm`.
- Tworzenie i sprawdzanie sum kontrolnych SHA-256.

## Wymagania

- Windows 10 lub nowszy, 64-bit.
- NVDA jest opcjonalny, ale program ma dodatkowe komunikaty dla NVDA.
- Do obsługi archiwów wymagany jest 7-Zip z `7z.exe` w standardowej lokalizacji albo w `PATH`.

Instalator zawiera aplikację i bibliotekę `nvdaControllerClient.dll`. 7-Zip nie jest dołączony do instalatora.

## Instalacja

Pobierz najnowszy plik `AmigaFM-Setup-*.exe` z sekcji Releases i uruchom instalator. Program jest instalowany w `Program Files`, więc instalator poprosi o uprawnienia administratora. Instalator tworzy skrót w menu Start i ikonę na pulpicie.

## Skróty klawiaturowe

- `Tab`: przełącza panel.
- `Enter`: otwiera plik, katalog, archiwum lub zasób.
- `Backspace`: katalog nadrzędny.
- `Space`: zaznacza lub odznacza bieżący element.
- `Ctrl+A`: zaznacza wszystko.
- `Ctrl+C`, `Ctrl+X`, `Ctrl+V`: kopiuj, wytnij, wklej.
- `F2`: zmiana nazwy.
- `F7`: nowy katalog.
- `Delete`: usuwanie.
- `Ctrl+F`: wyszukiwanie.
- `Shift+F10` albo `Ctrl+M`: menu kontekstowe.
- `Alt`: menu główne.
- `Page Up` i `Page Down`: pierwszy i ostatni element listy.
- `Escape`: wyjście z wyników wyszukiwania albo zamknięcie dialogu.
- `Alt+F4`: zamknięcie programu z potwierdzeniem.

## Budowanie ze źródeł

```powershell
cargo build --release
```

Wynik znajduje się w `target\release\amiga_fm.exe`. Skrypt budowania kopiuje `x64\nvdaControllerClient.dll` do katalogu release.

## Budowanie instalatora

Wymagany jest Inno Setup 6.

```powershell
.\scripts\build-release.ps1
```

Instalator zostanie utworzony w `installer\output`.

## Licencje i zależności

`nvdaControllerClient.dll` jest dystrybuowany na licencji LGPL 2.1. Pełna treść licencji znajduje się w `license.txt`.

Obsługa archiwów korzysta z zewnętrznego programu 7-Zip zainstalowanego w systemie użytkownika.
