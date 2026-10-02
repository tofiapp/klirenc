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

1. Chcete-li pokračovat v dřívějších poznámkách, klikněte na **Otevřít soubor…** a vyberte `.txt`. Jeho název se zobrazí v poli **Název souboru** a nové zápisy se přidávají na konec.
2. V Excelu označte celý rozsah řádků přes všechny sloupce, stiskněte Ctrl+C a vložte ho do pole vlevo nahoře (Ctrl+V). Aplikace si ponechá jen **první sloupec** (*Údaj k ověření*) a **poslední sloupec** (*poznámka*), ostatní se zahodí hned při vložení.
3. Klikněte na **Vytvořit seznam**. Prázdné řádky se přeskočí. Seznam se nevytvoří a dosavadní seznam zůstane beze změny, když některý řádek má jen jeden sloupec nebo prázdný první sloupec.
4. Místo vstupního pole se zobrazí seznam s číslem a stavem každého řádku: **✓** ponecháno, **✗** vymazáno (zapsáno v poznámkách), **?** vrátit se později, bez značky ještě nerozhodnuto. Aktuální řádek je zvýrazněný, jeho hodnota je velkým písmem uprostřed a je vždy zkopírovaná ve schránce.
   - **Ponechat** označí řádek ✓.
   - **Vymazat** zapíše poznámku a označí řádek ✗. Prázdnou poznámku nezapíše, jen upozorní.
   - **? Vrátit se později** označí řádek ?. Takové řádky přijdou na řadu, až nezbude žádný nerozhodnutý.
   - Po rozhodnutí se automaticky přejde na další nerozhodnutý řádek.
   - Šipkami **▲ ▼**, kliknutím na řádek v seznamu nebo klávesami ↑ ↓ v seznamu se můžete posouvat ručně a rozhodnutí změnit. Změna z ✓ na Vymazat zapíše poznámku. Změna z ✗ na Ponechat poznámku odebere.
5. **+ Vložit mezi** přidá nový záznam hned za aktuální řádek. Aktuální řádek se nezmění.
6. **Zrušit seznam** (s potvrzením) smaže seznam a vstupní pole. Poznámky zůstanou.
7. Ukládání (vždy UTF-8 `.txt`, v aplikaci se nic nemaže):
   - **Uložit** zapíše do otevřeného souboru. Pokud jste v poli **Název souboru** změnili název, soubor se ve stejné složce **přejmenuje**.
   - **Uložit jako…** se zeptá na nové umístění. Původní soubor zůstane.
