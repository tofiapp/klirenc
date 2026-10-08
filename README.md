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

1. V Excelu otevřete soubor **_kontrola_clearance_v4** uložený v počítači (ne online verzi) a makrem načtěte data. Data musí být na listu **ZJISTENI**, 1. řádek je záhlaví.
2. V aplikaci klikněte na **Načíst z Excelu**. Aplikace si přečte data přímo z otevřeného Excelu a použije tyto sloupce:
   - **1.** číslo revize (zobrazí se nad seznamem jako „Revize …“),
   - **3.** složka (datum RRMMDD),
   - **9.** stanice a **10.** délka (výběry v záhlaví složky),
   - **12.** údaj k ověření,
   - **16.** odkaz na mapu (dmwmap://…),
   - **36.** poznámka při Vymazat.

   Záloha: řádky lze také zkopírovat z Excelu, vložit do pole (Ctrl+V) a kliknout na **Vytvořit seznam**.
3. Záznamy jsou roztříděné do **složek** podle data od nejstarší. Kliknutím na záhlaví složky ji sbalíte nebo rozbalíte. V záhlaví jsou dva výběry, **stanice** a **délka** (včetně „Vše“). Ostatní záznamy složky se schovají a šipky i Ponechat / Vymazat procházejí jen zobrazené záznamy.
4. Aktuální hodnota se zobrazí velkým písmem a je vždy zkopírovaná ve schránce. Značky v seznamu: **✓** ponecháno, **✗** vymazáno (zapsáno v poznámkách), **?** vrátit se později.
   - **Ponechat** označí řádek ✓, **Vymazat** zapíše poznámku a označí řádek ✗, **? Vrátit se později** označí řádek ?.
   - Po rozhodnutí se přejde na další nerozhodnutý zobrazený řádek.
   - Šipkami **▲ ▼**, kliknutím na řádek nebo klávesami ↑ ↓ se můžete posouvat ručně a rozhodnutí změnit.
   - **Zobrazit na mapě** otevře odkaz aktuálního záznamu.
5. **Vkládání do jiné aplikace:** dokud aplikace běží, **Ctrl + levé kliknutí** do pole v jiném okně vloží aktuální údaj (obsah pole se nahradí). Vypnout to jde zaškrtávátkem nahoře.
6. **Zrušit seznam** (s potvrzením) smaže seznam a vstupní pole. Poznámky zůstanou.
7. Ukládání poznámek (vždy UTF-8 `.txt`):
   - **Otevřít soubor…** načte dřívější poznámky, nové zápisy se přidávají na konec.
   - **Uložit** zapíše do otevřeného souboru. Pokud jste v poli **Název souboru** změnili název, soubor se ve stejné složce přejmenuje.
   - **Uložit jako…** se zeptá na nové umístění.
