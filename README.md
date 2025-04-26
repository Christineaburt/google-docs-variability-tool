# ✏️ Variability Scoring Tool for Google Docs (Beta)

This **beta version** of a Google Apps Script-based tool evaluates text variation within structured Google Docs tables. It was **reverse-engineered based on observed behavior of Automated Insights’ Wordsmith variability scoring tool** — created to support editorial teams testing multiple content variations across regions, audiences, or SEO experiments.

---

## 🎯 Purpose

Content editors often generate multiple rewrites of the same sentence or concept to improve clarity, tone, or performance. This tool helps **quantify how diverse those rewrites actually are** by comparing lexical and structural differences, offering a fast, repeatable way to audit editorial variability at scale.

---

## 🔧 Features

- ✅ Extracts variations from Google Docs table rows or columns
- ✅ Tokenizes and normalizes sentence structures
- ✅ Scores each set of variations for diversity using a simplified Jaccard approach
- ✅ Logs results for editor review
- ✅ Built in pure Google Apps Script — no add-ons, no external dependencies

---

## ⚙️ How to Use (Google Apps Script)

1. **Create or open a Google Doc**
   - Include a table where each row or column contains alternate phrasings of the same base content

2. **Open [script.google.com](https://script.google.com)**
   - Create a new Apps Script project (can be container-bound or standalone)
   - Paste in the contents of `variabilityTool.gs`

3. **Run the script**
   - Authorize permissions
   - The tool will:
     - Read content from the Doc's table(s)
     - Tokenize and normalize the text
     - Score each variation set

4. **View results**
   - Scores and token-level breakdowns are logged or returned
   - High score = higher diversity among entries

---

## 🧪 Example Input Table

| Variation 1                  | Variation 2                  | Variation 3                  |
|-----------------------------|------------------------------|------------------------------|
| Florida drivers save $234   | Florida motorists save $234  | Florida drivers get $234 off |

🧠 Output (simplified):

"variabilityScore": 32

🧪 Beta Notice: This tool is currently in beta. The final version will include expanded cloud integration using the Google Cloud SDK and APIs to automate variability scoring within Google Docs.
