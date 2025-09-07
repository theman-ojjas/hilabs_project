---

# HiLabs Hackathon 2025 – Free-Text Roster Email Parser

## 📌 Problem Statement

This project is a solution to the **HiLabs Hackathon 2025** challenge:
Parsing and standardizing **free-text roster emails** into a clean, structured **Excel output** that exactly matches the provided template.

The tool:

* Reads input emails in **`.eml` format**.
* Extracts critical roster information (Provider Name, NPI, TIN, Specialty, Effective Date, etc.).
* Fills missing values with **“Information not found”**.
* Outputs a standardized Excel file (`Output.xlsx`) that matches the **Output Format.xlsx** specification.

---

## ⚙️ Setup Instructions

### 1. Clone the Repository

```bash
git clone <your-repo-link>
cd <repo-folder>
```

### 2. Install Requirements

Ensure you have **Python 3.10+** installed.
Install dependencies:

```bash
pip install -r requirements.txt
```

`requirements.txt` should include:

```
pandas
openpyxl
ollama
```

---

### 3. Setup Ollama & Model

This solution uses **Ollama** with a custom local model for parsing.
We used granite3.3:8b in Ollama because it offers a strong balance between accuracy and efficiency for parsing unstructured free-text emails. It runs fully locally (no third-party APIs), satisfying hackathon rules, and is well-suited for information extraction tasks like identifying provider details (NPI, TIN, dates, specialties). This makes it reliable, fast, and compliant for the roster parsing challenge.

1. Install [Ollama](https://ollama.ai).
2. Pull the base model:

   ```bash
   ollama run granite3.3:8b
   ```
3. Create the custom model alias (`mario`):

   ```bash
   ollama create mario -f Modelfile
   ```

---

### Output: `Output.xlsx`

| Transaction Type | Provider Name | Provider NPI | TIN       | Specialty  | Effective Date | ... |
| ---------------- | ------------- | ------------ | --------- | ---------- | -------------- | --- |
| Add              | John Doe      | 1234567890   | 458888885 | Cardiology | 09/01/2025     | ... |


