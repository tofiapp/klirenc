# klirenc

Lokální pomocník pro Windows (PowerShell + WinForms) pro ruční procházení řádků zkopírovaných z Excelu. Funguje offline, nic se neinstaluje a nepotřebuje admin práva. Data se na disk zapisují jen tehdy, když sami uložíte poznámky.

## Spuštění

**Dvojklikem na `spustit.vbs`** (musí být ve stejné složce jako `klirenc.ps1`). Otevře se jen okno aplikace, bez černého okna příkazového řádku.
`spustit.cmd` dělá totéž, jen na okamžik problikne příkazový řádek.
Dvojklik přímo na `klirenc.ps1` ho ve Windows jen otevře v Poznámkovém bloku.

Nebo z příkazového řádku ve složce se skriptem:

```
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\klirenc.ps1
```

Parametr `-ExecutionPolicy Bypass` platí jen pro tento jeden proces, systémovou politiku nemění.

## Použití

1. Chcete-li pokračovat v dřívějších poznámkách, klikněte na **Otevřít soubor…** a vyberte `.txt`. Nové zápisy se pak přidávají na jeho konec.
2. V Excelu označte celý rozsah řádků přes všechny sloupce (např. 25), stiskněte Ctrl+C a vložte ho do pole vlevo nahoře (Ctrl+V). Aplikace použije jen **první sloupec** (*Údaj k ověření*) a **poslední sloupec** (*Poznámka při Vymazat*). Sloupce mezi nimi se zahodí hned při vložení, takže v poli zůstanou jen tyto dva. Pod polem je vidět počet řádků a kolik sloupců mělo vložení.
3. Případně napište **Nadpis** a klikněte na **Vytvořit seznam**. Vyplněný nadpis se přidá na konec poznámek. Prázdné řádky se přeskočí. Seznam se nevytvoří a dosavadní seznam zůstane beze změny, když:
   - některý řádek má jen jeden sloupec,
   - řádky mají různý počet sloupců,
   - první sloupec je prázdný.
4. Po vytvoření se v horním poli zobrazí seznam. U každého řádku je vidět stav: **✓** ponecháno, **✗** vymazáno (zapsáno v poznámkách), bez značky ještě nerozhodnuto. Aktuální řádek je zvýrazněný. Jeho hodnota je velkým písmem uprostřed a je vždy zkopírovaná ve schránce.
   - **Ponechat** označí řádek ✓.
   - **Vymazat** zapíše poznámku (poslední sloupec) a označí řádek ✗. Prázdnou poznámku nezapíše, jen upozorní.
   - Po rozhodnutí se automaticky přejde na **další nerozhodnutý** řádek.
   - Šipkami **▲ ▼** nebo kliknutím na řádek v horním seznamu se můžete posouvat ručně a rozhodnutí změnit. Změna z ✓ na Vymazat zapíše poznámku. Změna z ✗ na Ponechat poznámku z poznámek odebere.
   - Dokud je seznam vytvořený, horní pole nejde upravovat. Nová data vložíte až po **Zrušit seznam**.
5. Další nadpis do stejných poznámek přidáte kdykoli tlačítkem **+ Přidat nadpis**.
6. **Vložit mezi** přidá nový záznam hned za aktuální řádek. Aktuální řádek se nezmění.
7. **Zrušit seznam** (s potvrzením) smaže seznam, pozici a vstupní pole. Nadpis a poznámky zůstanou.
8. **Uložit** uloží poznámky do otevřeného (nebo naposledy uloženého) souboru. **Uložit jako…** se zeptá na nový soubor. Ukládá se jako UTF-8 `.txt` a v aplikaci se nic nemaže.
