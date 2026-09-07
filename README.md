# Otvorené objednávky — manažérsky prehľad

Webový dashboard, ktorý si nahráte v prehliadači priamo xlsx súbor s otvorenými
objednávkami a on sám vytvorí:

- súhrnné kartičky (všetky otvorené vs. nad 45 dní, počet aj suma v €)
- stĺpcový graf podľa prevádzok, prepínateľný podľa počtu / podľa sumy,
  zoradený zostupne, s hodnotami pri stĺpcoch
- kompletnú tabuľku s riadkom Spolu
- (voliteľne) dva trendové grafy vývoja za celú skupinu od začiatku roka,
  ak súbor obsahuje záložku s mesačným vývojom

Všetko sa spracúva **iba v prehliadači** — súbor sa nikam neposiela, nič sa
neukladá na server.

## Očakávaná štruktúra xlsx súboru

- záložka, ktorej názov obsahuje slovo **PODKLAD** — musí mať dve pivot
  tabuľky vedľa seba, každá s hlavičkou stĺpca `Označenia riadkov` a
  nasledujúcimi stĺpcami súčtu sumy a počtu (presne ako v pôvodnom reporte)
- voliteľne záložka, ktorej názov obsahuje `nad 45 dní od…` — s riadkom
  `Spolu` a mesačnými pármi stĺpcov `Suma` / `Počet…`

Aplikácia si sama podľa veľkosti súčtu určí, ktorá pivot tabuľka je "všetky
otvorené objednávky" a ktorá "nad 45 dní", takže by mala fungovať aj na
budúcich mesačných verziách súboru bez úprav kódu.

## Spustenie lokálne

Vyžaduje [Node.js](https://nodejs.org/) (verzia 18 alebo novšia).

```bash
npm install
npm run dev
```

Otvorí sa na `http://localhost:5173`.

## Nahratie do GitHubu

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<vase-meno>/<nazov-repozitara>.git
git push -u origin main
```

## Nasadenie (GitHub Pages)

V repozitári je pripravený workflow `.github/workflows/deploy.yml`, ktorý pri
každom push na `main` automaticky zbuilduje aplikáciu a nasadí ju na GitHub
Pages.

Stačí v repozitári zapnúť: **Settings → Pages → Source: GitHub Actions**.
Po prvom úspešnom behu workflow bude aplikácia dostupná na
`https://<vase-meno>.github.io/<nazov-repozitara>/`.

### Alternatíva: manuálny deploy cez `gh-pages`

```bash
npm run deploy
```

(vyžaduje balík `gh-pages`, ktorý je už v `devDependencies`, a nastavený
vzdialený repozitár).

## Technológie

- [Vite](https://vitejs.dev/) + [React](https://react.dev/)
- [SheetJS (xlsx)](https://sheetjs.com/) — čítanie xlsx súboru v prehliadači
- [Recharts](https://recharts.org/) — grafy
- [lucide-react](https://lucide.dev/) — ikony
