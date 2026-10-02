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
2. V Excelu označte sloupec *Údaj k ověření*, stiskněte Ctrl+C a v levém poli klikněte na **Načíst ze schránky**. Totéž udělejte se sloupcem *Poznámka při NE* v pravém poli. Data můžete do polí vložit i ručně (Ctrl+V). Pod každým polem je vidět počet řádků. Když se počty liší, zčervenají.
3. Klikněte na **Vytvořit seznam**. Řádky se spárují podle pořadí a řádky prázdné v obou sloupcích se přeskočí. Seznam se nevytvoří a dosavadní seznam zůstane beze změny, když:
   - se počty řádků liší,
   - některý řádek má poznámku, ale prázdný údaj,
   - pole obsahuje víc sloupců.
4. Aktuální hodnota se zobrazí velkým písmem a automaticky se zkopíruje do schránky. Ve vstupních polích nahoře je zvýrazněný řádek, který je právě na řadě. Pod hodnotou je vidět, co se zapíše při NE.
   - **ANO** přejde na další řádek.
   - **NE** zapíše do poznámek druhý sloupec a přejde dál. Prázdnou poznámku nezapíše, jen upozorní.
5. **Vložit mezi** přidá nový záznam hned za aktuální řádek. Aktuální řádek se nezmění.
6. **Vymazat seznam** (s potvrzením) smaže seznam, pozici a importní vstup. Nadpis a poznámky zůstanou.
7. **Uložit poznámky jako…** uloží poznámky jako UTF-8 `.txt`. Nic se přitom nemaže.
