# Kontrola Clearance

Lokální pomocník pro Windows (PowerShell + WPF) pro ruční procházení řádků zkopírovaných z Excelu. Funguje offline, nic se neinstaluje a nepotřebuje admin práva. Data se na disk zapisují jen tehdy, když sami uložíte poznámky.

## Spuštění

**Dvojklikem na `spustit.vbs`** (musí být ve stejné složce jako `klirenc.ps1`). Otevře se jen okno aplikace, bez černého okna příkazového řádku.
`spustit.cmd` dělá totéž, jen na okamžik problikne příkazový řádek.
Dvojklik přímo na `klirenc.ps1` ho ve Windows jen otevře v Poznámkovém bloku.

**Zástupce s ikonou KC:** při každém spuštění aplikace se automaticky vytvoří nebo opraví zástupce **Kontrola Clearance** ve složce aplikace a na ploše. Vždy vede na aktuální složku. Pokud je zástupce připnutý na hlavním panelu, opraví se také. Číslo verze je vidět v titulku okna.

Nebo z příkazového řádku ve složce se skriptem:

```
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\klirenc.ps1
```

Parametr `-ExecutionPolicy Bypass` platí jen pro tento jeden proces, systémovou politiku nemění.

## Použití

1. Chcete-li pokračovat v dřívějších poznámkách, klikněte na **Otevřít soubor…** a vyberte `.txt`. Jeho název se zobrazí v poli **Název souboru** a nové zápisy se přidávají na konec.
2. V Excelu označte celý rozsah řádků přes všechny sloupce (34), stiskněte Ctrl+C a vložte ho do pole vlevo nahoře (Ctrl+V). Aplikace si ponechá jen 3 sloupce, ostatní zahodí hned při vložení:
   - **1. sloupec** = složka (kategorie),
   - **10. sloupec** = *Údaj k ověření*,
   - **poslední sloupec** = poznámka při Vymazat.
3. Klikněte na **Vytvořit seznam**. Záznamy se roztřídí do **složek** podle 1. sloupce. Složky jsou seřazené podle data RRMMDD od nejstarší, složky bez data jsou na konci. Kliknutím na záhlaví složky ji sbalíte nebo rozbalíte. U každé složky je vidět, kolik záznamů je hotovo.
4. Místo vstupního pole se zobrazí seznam s číslem a stavem každého řádku: **✓** ponecháno, **✗** vymazáno (zapsáno v poznámkách), **?** vrátit se později, bez značky ještě nerozhodnuto. Aktuální řádek je zvýrazněný, jeho hodnota je velkým písmem uprostřed a je vždy zkopírovaná ve schránce.
   - **Ponechat** označí řádek ✓.
   - **Vymazat** zapíše poznámku a označí řádek ✗. Prázdnou poznámku nezapíše, jen upozorní.
   - **? Vrátit se později** označí řádek ?. Takové řádky přijdou na řadu, až nezbude žádný nerozhodnutý.
   - Po rozhodnutí se automaticky přejde na další nerozhodnutý řádek.
   - Šipkami **▲ ▼**, kliknutím na řádek v seznamu nebo klávesami ↑ ↓ v seznamu se můžete posouvat ručně a rozhodnutí změnit. Změna z ✓ na Vymazat zapíše poznámku. Změna z ✗ na Ponechat poznámku odebere.
5. **Vkládání do jiné aplikace:** dokud aplikace běží (i na pozadí), **Ctrl + levé kliknutí** do pole v jiném okně vloží aktuální údaj (obsah pole se nahradí). Vypnout to jde zaškrtávátkem **Vkládat Ctrl + kliknutím** nahoře.
5b. **+ Vložit mezi** otevře okno, do kterého vložíte řádky z Excelu stejně jako do hlavního pole. Zařadí se hned za aktuální řádek, každý do své složky. Aktuální řádek se nezmění.
6. **Zrušit seznam** (s potvrzením) smaže seznam a vstupní pole. Poznámky zůstanou.
7. Ukládání (vždy UTF-8 `.txt`, v aplikaci se nic nemaže):
   - **Uložit** zapíše do otevřeného souboru. Pokud jste v poli **Název souboru** změnili název, soubor se ve stejné složce **přejmenuje**.
   - **Uložit jako…** se zeptá na nové umístění. Původní soubor zůstane.
