import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# Přístupové heslo
PASSWORD = "sakul7891"

password = st.text_input("Zadejte přístupové heslo", type="password")

if password != PASSWORD:
    st.stop()

# NOVÉ IMPORTY PRO DOCX
from io import BytesIO
from docx import Document
from docx.shared import Inches

st.set_page_config(page_title="Decision Engine – Půjčovna", layout="centered")

st.title("Decision Engine – Půjčovna malostrojů")
st.write("Interaktivní simulace zisku a rizika pro Mini Bagr vs. Pásový dumper.")

# --- Načtení dat ---
@st.cache_data
def load_data():
    return pd.read_excel("data/mvp_pujcovna_malostroju.xlsx")

df = load_data()

# --- Výběr typů ---
mini_bagry = df[df["type"] == "Mini Bagr"]
dumper = df[df["type"] == "Pásový dumper"]

if mini_bagry.empty or dumper.empty:
    st.error("V datech chybí typ 'Mini Bagr' nebo 'Pásový dumper'. Zkontroluj sloupec 'type'.")
    st.stop()

avg_daily_mini = float(mini_bagry["daily_cost_czk"].mean())
avg_daily_dumper = float(dumper["daily_cost_czk"].mean())

# --- Ovládání (slidery) ---
st.sidebar.header("Nastavení simulace")

fixed_cost = st.sidebar.number_input("Fixní náklad (Kč / měsíc)", min_value=0, value=120000, step=5000)

scenario = st.sidebar.selectbox(
    "Scénář trhu",
    ["Pesimistický", "Realistický", "Optimistický"]
)

if scenario == "Pesimistický":
    min_days, max_days = 8, 15
elif scenario == "Realistický":
    min_days, max_days = 15, 25
else:
    min_days, max_days = 22, 30

n = st.sidebar.slider("Počet simulací (měsíců)", 200, 5000, 1000, step=200)

expected_days = st.sidebar.slider("Odhad reálné poptávky (dny/měsíc)", 0, 31, 18)

# --- Simulace Monte Carlo ---
util_days = np.random.randint(min_days, max_days + 1, size=n)

rev_mini = avg_daily_mini * util_days
rev_dumper = avg_daily_dumper * util_days

profit_mini = rev_mini - fixed_cost
profit_dumper = rev_dumper - fixed_cost

# realistický scénář při očekávané poptávce
expected_profit_mini_real = avg_daily_mini * expected_days - fixed_cost
expected_profit_dumper_real = avg_daily_dumper * expected_days - fixed_cost

# riziko ztráty
prob_loss_mini = float(np.mean(profit_mini < 0))
prob_loss_dumper = float(np.mean(profit_dumper < 0))

# break-even
break_even_mini = fixed_cost / avg_daily_mini if avg_daily_mini > 0 else np.nan
break_even_dumper = fixed_cost / avg_daily_dumper if avg_daily_dumper > 0 else np.nan

# --- Porovnávací metriky (pro AI text) ---
expected_profit_mini = float(np.mean(profit_mini))
expected_profit_dumper = float(np.mean(profit_dumper))

risk_adv_pct = (prob_loss_mini - prob_loss_dumper) * 100

p10_mini = float(np.percentile(profit_mini, 10))
p10_dumper = float(np.percentile(profit_dumper, 10))
p50_mini = float(np.percentile(profit_mini, 50))
p50_dumper = float(np.percentile(profit_dumper, 50))
p90_mini = float(np.percentile(profit_mini, 90))
p90_dumper = float(np.percentile(profit_dumper, 90))

# --- Doporučení (jednoduché pravidlo) ---
if (prob_loss_dumper < prob_loss_mini) and (break_even_dumper < break_even_mini):
    recommendation = "✅ KUP DUMPER"
else:
    recommendation = "⚠️ NEKUPOVAT / DOPLNIT DATA"

# --- Výstupy nahoře ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("Mini Bagr")
    st.metric("Průměrná denní cena", f"{avg_daily_mini:,.0f} Kč")
    st.metric("Break-even", f"{break_even_mini:.2f} dní")
    st.metric("Riziko ztráty", f"{prob_loss_mini*100:.1f} %")

with col2:
    st.subheader("Pásový Dumper")
    st.metric("Průměrná denní cena", f"{avg_daily_dumper:,.0f} Kč")
    st.metric("Break-even", f"{break_even_dumper:.2f} dní")
    st.metric("Riziko ztráty", f"{prob_loss_dumper*100:.1f} %")

st.subheader("Doporučení")
st.write(recommendation)

# --- Realistický scénář ---
st.subheader("Realistický scénář při očekávané poptávce")

col3, col4 = st.columns(2)
with col3:
    st.metric("Mini Bagr – očekávaný zisk", f"{expected_profit_mini_real:,.0f} Kč")
with col4:
    st.metric("Dumper – očekávaný zisk", f"{expected_profit_dumper_real:,.0f} Kč")

# --- AI doporučení ---
st.subheader("AI doporučení (vysvětlení v lidské řeči)")

if "KUP DUMPER" in recommendation:
    ai_text = f"""
**Proč vychází lépe Dumper:**
- Očekávaný měsíční zisk: **Dumper {expected_profit_dumper:,.0f} Kč** vs. **Mini Bagr {expected_profit_mini:,.0f} Kč**
- Dumper má **o {risk_adv_pct:.1f} p.b. nižší riziko ztráty**
- Break-even: **Dumper {break_even_dumper:.2f} dní** vs. **Mini Bagr {break_even_mini:.2f} dní**

**Scénáře zisku (Kč) – aby bylo jasné riziko:**
- 10% pesimistický scénář: Dumper **{p10_dumper:,.0f}**, Mini Bagr **{p10_mini:,.0f}**
- Medián: Dumper **{p50_dumper:,.0f}**, Mini Bagr **{p50_mini:,.0f}**
- 90% optimistický scénář: Dumper **{p90_dumper:,.0f}**, Mini Bagr **{p90_mini:,.0f}**

**Co je potřeba doplnit před finálním nákupem:**
- reálnou poptávku po dumperu,
- servisní náklady / odstávky / pojištění,
- sezónnost (léto vs zima).
"""
else:
    ai_text = f"""
**Proč nejde dát jasné „kup“:**
- Rozdíl není dostatečně jednoznačný při zvolených parametrech.

**Rychlá fakta:**
- Očekávaný zisk Dumper: **{expected_profit_dumper:,.0f} Kč**
- Očekávaný zisk Mini Bagr: **{expected_profit_mini:,.0f} Kč**
- Riziko ztráty Dumper: **{prob_loss_dumper*100:.1f} %**
- Riziko ztráty Mini Bagr: **{prob_loss_mini*100:.1f} %**

**Další krok:**
Doplň data o poptávce, servisních nákladech a sezónnosti.
"""

st.markdown(ai_text)

# --- Graf 1: Distribuce zisku (Monte Carlo) ---
st.subheader("Porovnání distribuce zisku (Monte Carlo)")

fig_mc = plt.figure()
plt.hist(profit_mini, bins=25, alpha=0.5, label="Mini Bagr")
plt.hist(profit_dumper, bins=25, alpha=0.5, label="Dumper")
plt.axvline(0)
plt.xlabel("Zisk / ztráta (Kč)")
plt.ylabel("Počet simulovaných měsíců")
plt.title("Distribuce zisku")
plt.legend()
st.pyplot(fig_mc)

# --- Graf 2: Zisk vs vytížení (citlivost / break-even) ---
st.subheader("Citlivostní analýza: Zisk vs. vytížení (break-even)")

days_grid = np.arange(0, 32)  # 0..31 dní

profit_curve_mini = avg_daily_mini * days_grid - fixed_cost
profit_curve_dumper = avg_daily_dumper * days_grid - fixed_cost

fig_be = plt.figure()
plt.plot(days_grid, profit_curve_mini, label="Mini Bagr")
plt.plot(days_grid, profit_curve_dumper, label="Dumper")
plt.axhline(0)

plt.axvline(break_even_mini, linestyle="--")
plt.axvline(break_even_dumper, linestyle="--")

plt.xlabel("Vytížení (dny/měsíc)")
plt.ylabel("Zisk / ztráta (Kč)")
plt.title("Zisk podle vytížení (fixed_cost = zvolený fixní náklad)")
plt.legend()
st.pyplot(fig_be)

st.caption(
    f"Break-even: Mini Bagr ≈ {break_even_mini:.2f} dní, Dumper ≈ {break_even_dumper:.2f} dní. "
    "Nad těmito hodnotami je stroj v průměru v zisku."
)

# --- Malé shrnutí ---
st.subheader("Interpretace (laicky)")
st.write(
    f"- Mini Bagr má riziko ztráty **{prob_loss_mini*100:.1f} %** při zvolených parametrech.\n"
    f"- Dumper má riziko ztráty **{prob_loss_dumper*100:.1f} %**.\n"
    f"- Break-even říká, kolik dní v měsíci musí být stroj pronajatý, aby pokryl fixní náklady."
)

st.subheader("Stáhnout report (DOCX)")

def fig_to_png_bytes(fig):
    img = BytesIO()
    fig.savefig(img, format="png", dpi=200, bbox_inches="tight")
    img.seek(0)
    return img

def build_docx_report():
    doc = Document()

    # Titulek
    doc.add_heading("Decision Engine – Investiční analýza", level=1)
    doc.add_paragraph("Porovnání: Mini Bagr vs. Pásový dumper (Monte Carlo simulace).")

    # Shrnutí
    doc.add_heading("Shrnutí výsledků", level=2)
    doc.add_paragraph(f"Fixní náklad (Kč/měsíc): {fixed_cost:,.0f}")
    doc.add_paragraph(f"Rozsah vytížení: {min_days}–{max_days} dní/měsíc")
    doc.add_paragraph(f"Počet simulací: {n:,}")
    doc.add_paragraph(f"Odhad reálné poptávky: {expected_days} dní/měsíc")

    # Tabulka metrik
    doc.add_heading("Klíčové metriky", level=2)
    table = doc.add_table(rows=1, cols=3)
    hdr = table.rows[0].cells
    hdr[0].text = "Metrika"
    hdr[1].text = "Mini Bagr"
    hdr[2].text = "Pásový dumper"

    rows = [
        ("Průměrná denní cena", f"{avg_daily_mini:,.0f} Kč", f"{avg_daily_dumper:,.0f} Kč"),
        ("Break-even", f"{break_even_mini:.2f} dní", f"{break_even_dumper:.2f} dní"),
        ("Riziko ztráty", f"{prob_loss_mini*100:.1f} %", f"{prob_loss_dumper*100:.1f} %"),
        ("Oček. zisk (Monte Carlo)", f"{expected_profit_mini:,.0f} Kč", f"{expected_profit_dumper:,.0f} Kč"),
        ("P10 (pesimisticky)", f"{p10_mini:,.0f} Kč", f"{p10_dumper:,.0f} Kč"),
        ("Medián (P50)", f"{p50_mini:,.0f} Kč", f"{p50_dumper:,.0f} Kč"),
        ("P90 (optimisticky)", f"{p90_mini:,.0f} Kč", f"{p90_dumper:,.0f} Kč"),
        ("Realistický zisk (poptávka)", f"{expected_profit_mini_real:,.0f} Kč", f"{expected_profit_dumper_real:,.0f} Kč"),
    ]

    for r in rows:
        row_cells = table.add_row().cells
        row_cells[0].text = r[0]
        row_cells[1].text = r[1]
        row_cells[2].text = r[2]

    # Doporučení
    doc.add_heading("Doporučení", level=2)
    doc.add_paragraph(recommendation)

    # AI text (bez markdown formátování)
    doc.add_heading("AI vysvětlení (laicky)", level=2)
    doc.add_paragraph(ai_text.replace("**", "").replace("- ", "• "))

    # Grafy
    doc.add_heading("Grafy", level=2)

    doc.add_paragraph("1) Distribuce zisku (Monte Carlo)")
    img1 = fig_to_png_bytes(fig_mc)
    doc.add_picture(img1, width=Inches(6.0))

    doc.add_paragraph("2) Zisk vs. vytížení (break-even)")
    img2 = fig_to_png_bytes(fig_be)
    doc.add_picture(img2, width=Inches(6.0))

    # Výstup do paměti
    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out

docx_file = build_docx_report()

st.download_button(
    label="Stáhnout report (DOCX)",
    data=docx_file,
    file_name="decision_report.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)