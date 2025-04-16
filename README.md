# Troglodyte

A helpful script set for export various bulk data from office documents in **pdf**.
Sada skriptů pro export různých objemných dat z úředních dokumentů ve formátu **pdf**.

## Changelog

- 31.03.2025 **01.001**: The first test.
- 01.04.2025 **01.002**: Directory lister in library, US->CZ number format conversion, code cleanup.
- 16.04.2025 **02.001**: Added CSBI from BMW, common output file, function for using it, optimized previous code.

## Documentation

### Scripts and Library modules

- **CreditAdvicePorscheA.py** - Extract data form **Credit Advice** document form Porsche
- **CsviBmw.py** - Extract data from **Corrected Self Billed Invoice** document from BMW
- **PdfExtractor.py** - A library of special functions

### Použití

1) `python CreditAdvicePorscheA.py` - spuštění skriptu s defaultním pdf souborem nastaveným napevno v kódu (pouze pro testovací účely)

2) `python CreditAdvicePorscheA.py file.pdf` - spuštění skriptu se souborem file.pdf

3) `python CreditAdvicePorscheA.py directory` - spuštění skriptu s se všemi pdf soubory ve složce directory

- Vstupní soubor musí mít vždy příponu ***.pdf** nebo ***.PDF**
- Data jsou uložena do souboru ***.csv** stejného jména, jako měl vstupní soubor.

4) Stejné postupy platí pro CsbiBmw.py, výstup ze všech zpracovaných souborů se ukládá do společného souboru **data/bmw/output.csv**.