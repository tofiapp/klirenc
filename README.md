# klirenc

Lokální pomocník pro Windows (PowerShell + WinForms) pro ruční procházení řádků zkopírovaných z Excelu. Funguje offline, nic se neinstaluje a nepotřebuje admin práva. Data se na disk zapisují jen tehdy, když sami uložíte poznámky.

## Spuštění

**Dvojklikem na `spustit.cmd`** (musí být ve stejné složce jako `klirenc.ps1`).
Dvojklik přímo na `klirenc.ps1` ho ve Windows jen otevře v Poznámkovém bloku.

Nebo z příkazového řádku ve složce se skriptem:

```
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\klirenc.ps1
```

Parametr `-ExecutionPolicy Bypass` platí jen pro tento jeden proces, systémovou politiku nemění.

## Použití

1. Vyplňte **Nadpis poznámek** (volitelné).
2. V Excelu označte dva sloupce (*Údaj k ověření*, *Poznámka při NE*) a stiskněte Ctrl+C.
3. Klikněte na **Načíst ze schránky**. Prázdné řádky se přeskočí. Když některý řádek nemá přesně dva sloupce, nenačte se nic a dosavadní seznam zůstane beze změny.
4. Aktuální hodnota se zobrazí velkým písmem a automaticky se zkopíruje do schránky.
   - **ANO** přejde na další řádek.
   - **NE** zapíše do poznámek druhý sloupec a přejde dál. Prázdnou poznámku nezapíše, jen upozorní.
5. **Vložit mezi** přidá nový záznam hned za aktuální řádek. Aktuální řádek se nezmění.
6. **Vymazat seznam** (s potvrzením) smaže seznam, pozici a importní vstup. Nadpis a poznámky zůstanou.
7. **Uložit poznámky jako…** uloží poznámky jako UTF-8 `.txt`. Nic se přitom nemaže.
