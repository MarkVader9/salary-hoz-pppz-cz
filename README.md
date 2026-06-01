# SALARY s.r.o. - Rozhraní pro české zdravotní pojišťovny (XSD/XML)

<!-- ISO 9001:2015 Referenční číslo dokumentu: REF-IT-XSD-01 | Verze: 1.0.0 | Stav: SCHVÁLENO -->

Hlavní repozitář obsahující oficiální XSD schémata a XML šablony pro zákonný reporting zdravotním pojišťovnám ČR. Vyvíjeno a udržováno společností **SALARY s.r.o.** jako nedílná součást naší validační platformy pro mzdový core software.

## 📌 Přehled projektu

Tento repozitář slouží jako jediný zdroj pravdy (SSOT) pro formáty validace dat vyžadované při podnikovém zpracování mezd. Zajišťuje přísný soulad s legislativními požadavky ČR a **Standardy řízení kvality ISO 9001:2015** (konkrétně článek 8.1 – Plánování a řízení provozu).

### Spravovaná dokumentová rozhraní:
*   **HOZ** (Hromadné oznámení zaměstnavatele) – Registrace zaměstnanců, změny zdravotních pojišťoven a změny kategorií (Zákonná lhůta: 8 dnů).
*   **PPPZ** (Přehled o platbě pojistného zaměstnavatele) – Měsíční přehled o platbě pojistného, vyměřovací základy a prohlášení zaměstnavatele (Zákonná lhůta: do 20. dne následujícího měsíce).

## 🗂 Struktura repozitáře

```text
├── .github/              # CI/CD pipelines pro automatizovanou validaci XML
├── schemas/              # Oficiální XSD soubory
│   ├── hoz/
│   │   └── hoz_v2026.xsd # Schéma pro Hromadné oznámení zaměstnavatele
│   └── pppz/
│       └── pppz_v2026.xsd # Schéma pro Přehled o platbě pojistného
├── templates/            # Ukázky produkčních XML dat (mock data)
│   ├── hoz_sample.xml
│   └── pppz_sample.xml
└── README.md             # Tato dokumentace
```

## ⚖️ Právní a regulační rámec

Všechna schémata v tomto repozitáři jsou v souladu s platnou legislativou České republiky:
*   **Zákon č. 48/1997 Sb.** o veřejném zdravotním pojištění
*   **Zákon č. 592/1992 Sb.** o pojistném na veřejné zdravotní pojištění
*   **Zákon č. 300/2008 Sb.** o elektronických úkonech a autorizované konverzi dokumentů (přenosový protokol datových schránek)

## 🛠 Integrace a CI/CD validace

Každý commit do větve `main` spouští automatizované workflow, které ověřuje kompatibilitu a správnost schémat.

### Rychlá manuální validace (CLI)
Pro lokální validaci exportovaného XML souboru vůči našemu produkčnímu XSD schématu spusťte:
```bash
xmllint --schema schemas/hoz/hoz_v2026.xsd templates/hoz_sample.xml --noout
```

## 🔒 Bezpečnost a ochrana osobních údajů (GDPR)
*   **Politika nulového výskytu PII:** V tomto repozitáři je přísně zakázáno uvádět reálné osobní identifikační údaje (PII).
*   Všechny datové soubory v adresáři `templates/` používají náhodně generované testovací proměnné.
*   Přístupové tokeny a certifikáty elektronických podpisů musí být konfigurovány výhradně přes tajná prostředí (`GH_SECRETS`) a nikdy nesmí být pevně zapsány v kódu.

## 👥 Autoři a řízení změn

*   **Správce:** Vývojový tým SALARY s.r.o. (<dev@salary.cz>)
*   **Vlastník procesu:** Manažer pro compliance a řízení kvality (QA)
*   **ISO 9001 Auditní stopa:** Jakákoliv změna definic `.xsd` vyžaduje formální zápis do protokolu o změnovém řízení (Change Request).

---
© 2026 SALARY s.r.o. Všechna práva vyhrazena. Interní technická dokumentace.
