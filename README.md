# Decision Engine – Rental Equipment

## Project goal
The goal of this project is to help a rental company decide whether it should invest in a Mini Excavator or a Track Dumper.

## Method
The project uses:
- Data analysis (Pandas)
- Monte Carlo simulation
- Profit distribution analysis
- Break-even analysis
- Risk analysis

## Results
Simulation shows that the Track Dumper has:

- higher expected profit
- lower probability of loss

## Output
The program automatically generates:

- profit comparison graph
- automated investment report

## Technologies
Python
Pandas
NumPy
Matplotlib

## Spuštění lokálně

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Nastavení přístupového hesla (Streamlit secrets)

Aplikace je chráněna heslem, které se načítá výhradně ze Streamlit secrets
(`st.secrets["password"]`). Heslo nikdy nepatří do zdrojového kódu ani do Gitu.

**Lokálně**

1. Zkopírujte šablonu: `cp .streamlit/secrets.toml.example .streamlit/secrets.toml`
2. V souboru `.streamlit/secrets.toml` nastavte vlastní heslo.
3. Soubor `secrets.toml` je uvedený v `.gitignore`, takže se do repozitáře nedostane.

**Streamlit Community Cloud**

1. Otevřete aplikaci na share.streamlit.io → ⋮ → *Settings* → *Secrets*.
2. Vložte řádek `password = "..."` s vlastním heslem a uložte.

## Data

Repozitář obsahuje pouze fiktivní demonstrační data (`data/`).
