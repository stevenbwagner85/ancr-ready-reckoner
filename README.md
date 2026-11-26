# AnCR ready reckoner

A structured aid for judging analytical confidence. It turns an analyst's answers to a set of prompt questions into the ingredients of an Analytical Confidence Rating (AnCR) statement, one key judgement at a time.

It replaces an Excel workbook that relied on ActiveX scroll bars, which only run in desktop Excel for Windows and are now blocked by default in Microsoft 365.

## What is here

| File | What it is |
|---|---|
| `ancr-ready-reckoner.html` | The tool as a single file with draggable sliders. Save it to a local or shared drive and open it in a browser. It works offline, installs nothing, and saves its results as an Excel workbook. |
| `AnCR-ready-reckoner-template.xlsx` | The same assessment as a plain Excel workbook. No macros, no ActiveX. Scores are chosen from a list and drawn as a bar with conditional formatting. |
| `index.html` | The download page. |

A workbook saved from the HTML tool has exactly the same layout as the template, because both are produced by the same code. It can be edited in Excel and opened again in the HTML tool.

## How the assessment works

- Up to ten key judgements, each on its own tab or sheet.
- Nine factors under three criteria:
  1. **Source base**: coverage, consistency, source reliability, background knowledge.
  2. **Analysis**: argumentation, assumptions, challenge.
  3. **Complexity and volatility**: complexity, volatility. A higher score means a simpler, more stable subject.
- Each factor is scored from 0 to 100 in steps of 10 between two anchor statements, with a notes box for the reasoning.
- The analyst sets an overall score for each criterion and for the key judgement as a whole. **Nothing is averaged or calculated.**
- An overall score of 0 to 30 reads as Low, 40 to 60 as Moderate and 70 to 100 as High.
- Each key judgement ends with the ingredients for the statement: the ratings, the strongest and weakest factors, and a writing frame. The analyst writes the statement.
- An optional protective marking appears at the top and bottom of every tab and sheet.

## Using the HTML tool

1. Download `ancr-ready-reckoner.html` and save it wherever you like.
2. Double-click it. It opens in your browser.
3. Choose **Save as Excel** to keep a workbook. In Chrome and Edge you choose where it goes; other browsers put it in the downloads folder.
4. Choose **Open workbook** to carry on from a saved workbook.

The tool sends nothing over the network. It keeps work in progress in the browser's local storage on that computer; **Start again** removes it. If the browser blocks local storage, the tool says so and warns before the page is closed.

## Using the Excel template

Open it and complete the cells outlined in blue. The sheets are protected without a password so the layout stays intact; use Review > Unprotect Sheet to adapt it.

The protective marking is typed once on the Summary sheet. It is shown at the top and bottom of each sheet and repeats at the top of every printed page.

## Building from source

The generated files are committed, so there is nothing to build before use. To rebuild after changing anything in `src/`:

```
python3 build.py
python3 tests/run.py
```

The build needs Python 3 and either Node.js or macOS (it falls back to the JavaScriptCore shell that ships with macOS). There are no other dependencies.

| Source | Purpose |
|---|---|
| `src/model.js` | Criteria, factors, wording, rating bands, writing frame. Edit the wording here. |
| `src/workbook.js` | Builds the workbook and reads one back. |
| `src/zip.js` | Minimal ZIP reader and writer, so the tool needs no libraries. |
| `src/app.html`, `src/app.css`, `src/app.js` | The HTML tool's interface. |

`tests/run.py` builds a filled-in workbook, checks every part is well-formed XML, and checks the reader recovers the same data from it, from a re-compressed copy, and from a copy saved by desktop Excel.

## Branding

The Brunel University of London logo and the College of Arts, Law and Social Sciences name are the property of Brunel University of London and are not covered by any licence applied to the code.
