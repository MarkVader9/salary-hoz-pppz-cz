SALARY s.r.o. - Rozhraní pro české zdravotní pojišťovny (XSD/XML)

Hlavní repozitář obsahující oficiální schémata XSD a šablony XML pro regulační reporting pro české zdravotní pojišťovny (Zdravotní pojišťovny ČR). Vyvinuto a spravováno společností SALARY s.r.o. jako nedílná součást naší platformy pro validaci základního softwaru pro mzdy.
📌 Přehled projektu

Toto repozitář slouží jako jediný zdroj pravdy (SSOT) pro formáty validace dat vyžadované firemním zpracováním mezd. Zajišťuje přísné dodržování českých zákonných požadavků a standardů managementu kvality ISO 9001:2015 (konkrétně článku 8.1 - Provozní plánování a řízení).
Rozhraní pro spravované dokumenty:

HOZ (Hromadné oznámení zaměstnanců) - Registrace zaměstnanců, změny zdravotních pojišťoven a úpravy kategorií (Zákonný limit: 8 dní).
PPPZ (Přehled o platbě pojistného zaměstnavatele) - Měsíční úhrn pojistného, ​​výpočtové základy a výkazy zaměstnavatele (zákonný limit: 20. den následujícího měsíce).

🗂 Struktura úložiště

├── .github/ # kanály CI/CD pro automatizovanou validaci XML
├── schémata/ # Oficiální soubory XSD
│ ├── hoz/
│ │ └── hoz_v2026.xsd # Schéma Hromadného oznámení zaměstnavatele
│ └── pppz/
│ └── pppz_v2026.xsd # Schéma pro Přehled o platbě pojistného
├── šablony/ # Vzorky užitečného zatížení XML připravené na výrobu
│ ├── hoz_sample.xml
│ └── pppz_sample.xml
└── README.md # Tato dokumentace

⚖️ Právní a regulační rámec

Všechna schémata v tomto repozitáři splňují platné právní předpisy České republiky:

Zákon č. 48/1997 Sb. o veřejném zdravotním pojištění
Zákon č. 592/1992 Sb. o pojistném na veřejné zdravotní pojištění
Zákon č. 300/2008 Sb. o elektronických úkonech a autorizované konverzi dokumentů (protokol pro doručení datových schránek)

🛠 Integrace a validace CI/CD

Každý commit do hlavní větve spustí automatizovaný pracovní postup, který provede kontroly shody schématu.
Rychlé ruční ověření (CLI)

Chcete-li lokálně ověřit export XML oproti našemu produkčnímu XSD, spusťte:

xmllint --schema schemas/hoz/hoz_v2026.xsd templates/hoz_sample.xml --noout

🔒 Zabezpečení a ochrana osobních údajů (GDPR)

Zásady nulových PII: Toto repozitář přísně zakazuje zahrnutí skutečných osobních identifikovatelných údajů (PII).
Všechny datové soubory v šablonách/ používají náhodně generované falešné proměnné.
Přístupové tokeny a certifikáty elektronického podpisu musí být konfigurovány pomocí tajných klíčů prostředí (GH_SECRETS) a nikdy nemusejí být pevně kódovány.

👥 Autoři a správa změn

Správce: Vývojový tým SALARY s.r.o. (dev@salary.cz)
Vlastník procesu: Manažer pro dodržování předpisů a zajištění kvality
Auditní záznam ISO 9001: Jakákoli změna definic .xsd vyžaduje formální protokol žádosti o změnu (CR).

© 2026 SALARY s.r.o. Všechna práva vyhrazena. Interní technická dokumentace.
