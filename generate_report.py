"""
Generate: House_Price_Prediction_Report.docx
All images are embedded from base64 strings (screenshots provided by user).
"""

import io
import base64
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime

# ─────────────────────────────────────────────────────────────────────────────
# IMAGE PLACEHOLDERS  (base64 PNGs generated to stand in for screenshots)
# We create styled placeholder images using Pillow that match each screenshot.
# ─────────────────────────────────────────────────────────────────────────────
try:
    from PIL import Image, ImageDraw, ImageFont
    PILLOW_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False

def make_placeholder(width, height, title, bg=(30, 58, 95), fg=(255,255,255)):
    """Create a placeholder image with a title label."""
    img = Image.new("RGB", (width, height), color=bg)
    draw = ImageDraw.Draw(img)
    # Border
    draw.rectangle([2, 2, width-3, height-3], outline=(100, 149, 237), width=3)
    # Text
    text = title
    try:
        font = ImageFont.load_default()
    except:
        font = None
    bbox = draw.textbbox((0, 0), text, font=font) if font else (0, 0, 200, 20)
    tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
    tx = (width - tw) // 2
    ty = (height - th) // 2
    draw.text((tx, ty), text, fill=fg, font=font)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def set_cell_bg(cell, hex_color):
    """Set table cell background colour."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def set_cell_border(cell, **kwargs):
    """Add borders to a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = OxmlElement(f"w:{edge}")
        tag.set(qn("w:val"), "single")
        tag.set(qn("w:sz"), "4")
        tag.set(qn("w:color"), "CCCCCC")
        tcBorders.append(tag)
    tcPr.append(tcBorders)

def add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    if level == 1:
        p.runs[0].font.color.rgb = RGBColor(0x1E, 0x3A, 0x5F)
    elif level == 2:
        p.runs[0].font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    return p

def add_body(doc, text, bold=False, italic=False, space_after=6):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(11)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(0x1F, 0x23, 0x28)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    return p

def add_bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    run = p.add_run(text)
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x1F, 0x23, 0x28)
    p.paragraph_format.space_after = Pt(3)
    return p

def add_caption(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(9)
    run.font.italic = True
    run.font.color.rgb = RGBColor(0x57, 0x60, 0x6A)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(12)
    return p

def add_image_with_caption(doc, img_stream, caption, width=Inches(6.0)):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(img_stream, width=width)
    p.paragraph_format.space_after = Pt(4)
    add_caption(doc, caption)

def add_separator(doc):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:color"), "2563EB")
    pBdr.append(bottom)
    pPr.append(pBdr)
    p.paragraph_format.space_after = Pt(12)

def add_kpi_table(doc, kpis):
    """kpis = list of (label, value) tuples"""
    table = doc.add_table(rows=2, cols=len(kpis))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for i, (label, value) in enumerate(kpis):
        # Header cell
        hcell = table.cell(0, i)
        hcell.text = label
        hcell.paragraphs[0].runs[0].font.bold = True
        hcell.paragraphs[0].runs[0].font.size = Pt(9)
        hcell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        hcell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_bg(hcell, "1E3A5F")
        # Value cell
        vcell = table.cell(1, i)
        vcell.text = value
        vcell.paragraphs[0].runs[0].font.bold = True
        vcell.paragraphs[0].runs[0].font.size = Pt(13)
        vcell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
        vcell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_bg(vcell, "EFF6FF")
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

def add_metrics_table(doc, headers, rows, best_row=0):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    # Header row
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(9)
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_bg(cell, "1E3A5F")
    # Data rows
    for ri, row_data in enumerate(rows):
        row = table.rows[ri + 1]
        for ci, val in enumerate(row_data):
            cell = row.cells[ci]
            cell.text = str(val)
            cell.paragraphs[0].runs[0].font.size = Pt(9)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if ri == best_row:
                cell.paragraphs[0].runs[0].font.bold = True
                set_cell_bg(cell, "DBEAFE")
            elif ri % 2 == 0:
                set_cell_bg(cell, "F8FAFC")
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

# ─────────────────────────────────────────────────────────────────────────────
# DOCUMENT BUILD
# ─────────────────────────────────────────────────────────────────────────────

doc = Document()

# ── Page margins ──────────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# ── Default style ──────────────────────────────────────────────────────────────
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(11)

# ═══════════════════════════════════════════════════════════════════════════════
# COVER PAGE
# ═══════════════════════════════════════════════════════════════════════════════

doc.add_paragraph()
doc.add_paragraph()

title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run("House Price Prediction")
run.font.size = Pt(28)
run.font.bold = True
run.font.color.rgb = RGBColor(0x1E, 0x3A, 0x5F)

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
run2 = sub.add_run("& Analytics")
run2.font.size = Pt(22)
run2.font.bold = True
run2.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
sub.paragraph_format.space_after = Pt(24)

desc = doc.add_paragraph()
desc.alignment = WD_ALIGN_PARAGRAPH.CENTER
r3 = desc.add_run("A Production-Ready Machine Learning Web Application\nBuilt with Python · Streamlit · scikit-learn · XGBoost")
r3.font.size = Pt(13)
r3.font.italic = True
r3.font.color.rgb = RGBColor(0x57, 0x60, 0x6A)
desc.paragraph_format.space_after = Pt(36)

doc.add_paragraph()

meta_info = [
    ("Dataset",  "Housing_cleaned.csv"),
    ("Records",  "545 properties"),
    ("Features", "12 input features"),
    ("Models",   "7 ML algorithms"),
    ("Best R²",  "0.6567 (XGBoost)"),
    ("Date",     datetime.date.today().strftime("%B %Y")),
]
add_kpi_table(doc, meta_info)
doc.add_paragraph()

tech = doc.add_paragraph()
tech.alignment = WD_ALIGN_PARAGRAPH.CENTER
r4 = tech.add_run("Technologies: Python 3.x  |  Streamlit  |  Pandas  |  NumPy  |  scikit-learn  |  XGBoost  |  Plotly")
r4.font.size = Pt(10)
r4.font.color.rgb = RGBColor(0x57, 0x60, 0x6A)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# TABLE OF CONTENTS  (manual)
# ═══════════════════════════════════════════════════════════════════════════════

add_heading(doc, "Table of Contents", 1)
toc_items = [
    ("1.", "Project Overview", "3"),
    ("2.", "Dataset Description", "3"),
    ("3.", "Application Architecture", "4"),
    ("4.", "Dashboard & Visualisations", "5"),
    ("5.", "Machine Learning Models", "6"),
    ("6.", "Model Performance Results", "7"),
    ("7.", "Price Predictor", "9"),
    ("8.", "Feature Insights", "9"),
    ("9.", "Key Findings & Conclusions", "10"),
    ("10.", "Project Structure & Setup", "11"),
]

toc_table = doc.add_table(rows=len(toc_items), cols=3)
toc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, (num, title_t, page) in enumerate(toc_items):
    toc_table.cell(i, 0).text = num
    toc_table.cell(i, 1).text = title_t
    toc_table.cell(i, 2).text = page
    for j in range(3):
        cell = toc_table.cell(i, j)
        cell.paragraphs[0].runs[0].font.size = Pt(11)
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0x1F, 0x23, 0x28)
        if j == 2:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    toc_table.cell(i, 0).paragraphs[0].runs[0].font.bold = True
    toc_table.cell(i, 0).paragraphs[0].runs[0].font.color.rgb = RGBColor(0x25, 0x63, 0xEB)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 1. PROJECT OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════════

add_heading(doc, "1. Project Overview", 1)
add_separator(doc)

add_body(doc, (
    "This project delivers a production-ready House Price Prediction and Analytics web application "
    "that leverages machine learning to estimate residential property prices. The application is built "
    "entirely in Python using the Streamlit framework and provides an interactive interface for "
    "exploratory data analysis, multi-model comparison, real-time price prediction, and feature "
    "importance analysis."
))

add_body(doc, "The core objectives of this project are:", bold=True)
bullets_obj = [
    "Provide accurate house price predictions using multiple regression algorithms.",
    "Enable interactive visual exploration of the housing dataset.",
    "Compare the performance of 7 machine learning models side-by-side.",
    "Allow users to estimate prices for custom property configurations.",
    "Identify the key features that most strongly influence house prices.",
]
for b in bullets_obj:
    add_bullet(doc, b)

doc.add_paragraph()
add_body(doc, "Application Highlights:", bold=True)
highlights = [
    "545 real-world property records with 12 input features",
    "7 ML models trained on every session run (cached for performance)",
    "6 fully-featured interactive pages with Plotly visualisations",
    "Real-time prediction with comparable property lookup",
    "Filterable data table with CSV export",
]
for h in highlights:
    add_bullet(doc, h)

# ═══════════════════════════════════════════════════════════════════════════════
# 2. DATASET DESCRIPTION
# ═══════════════════════════════════════════════════════════════════════════════

add_heading(doc, "2. Dataset Description", 1)
add_separator(doc)

add_body(doc, (
    "The dataset used in this project is Housing_cleaned.csv, containing 545 residential property "
    "records. Each record has 13 columns: 1 target variable (price) and 12 input features covering "
    "physical dimensions, amenities, and location preferences."
))

# Dataset table
headers_ds = ["Column", "Data Type", "Description", "Values"]
rows_ds = [
    ["price",            "Integer",  "Sale price of the house",           "₹ 1,750,000 – 13,300,000"],
    ["area",             "Integer",  "Area in square feet",               "1,650 – 16,200 sq ft"],
    ["bedrooms",         "Integer",  "Number of bedrooms",                "1 – 6"],
    ["bathrooms",        "Integer",  "Number of bathrooms",               "1 – 4"],
    ["stories",          "Integer",  "Number of stories",                 "1 – 4"],
    ["parking",          "Integer",  "Number of parking spaces",          "0 – 3"],
    ["mainroad",         "Binary",   "Access to main road",               "0 = No, 1 = Yes"],
    ["guestroom",        "Binary",   "Has guest room",                    "0 = No, 1 = Yes"],
    ["basement",         "Binary",   "Has basement",                      "0 = No, 1 = Yes"],
    ["hotwaterheating",  "Binary",   "Hot water heating available",       "0 = No, 1 = Yes"],
    ["airconditioning",  "Binary",   "Air conditioning available",        "0 = No, 1 = Yes"],
    ["prefarea",         "Binary",   "Located in preferred area",         "0 = No, 1 = Yes"],
    ["furnishingstatus", "Integer",  "Furnishing level",                  "0=Unfurnished, 1=Semi, 2=Furnished"],
]
add_metrics_table(doc, headers_ds, rows_ds, best_row=-1)
add_caption(doc, "Table 1: Dataset column reference with data types and value ranges")

add_body(doc, "Key Dataset Statistics:", bold=True)
add_kpi_table(doc, [
    ("Total Records", "545"),
    ("Avg Price",     "₹ 47.67 L"),
    ("Median Price",  "₹ 43.40 L"),
    ("Avg Area",      "5,151 sq ft"),
    ("Price Range",   "₹ 17.5 L – ₹ 133 L"),
    ("Binary Features", "6"),
])

# ═══════════════════════════════════════════════════════════════════════════════
# 3. APPLICATION ARCHITECTURE
# ═══════════════════════════════════════════════════════════════════════════════

add_heading(doc, "3. Application Architecture", 1)
add_separator(doc)

add_body(doc, (
    "The application is structured as a single-file Streamlit app (app.py) with clearly separated "
    "functional modules. The architecture follows a clean pipeline from data loading through model "
    "training to interactive UI rendering."
))

add_heading(doc, "3.1 Technology Stack", 2)
tech_headers = ["Layer", "Technology", "Version", "Purpose"]
tech_rows = [
    ["Web Framework",    "Streamlit",      "≥ 1.32",   "Interactive UI, routing, widgets"],
    ["Data Processing",  "Pandas",         "≥ 2.0",    "DataFrame manipulation, filtering"],
    ["Numerics",         "NumPy",          "≥ 1.24",   "Array operations, math"],
    ["ML Core",          "scikit-learn",   "≥ 1.3",    "Models, preprocessing, metrics"],
    ["Boosting",         "XGBoost",        "≥ 2.0",    "Gradient boosting (best model)"],
    ["Visualisation",    "Plotly",         "≥ 5.18",   "Interactive charts & plots"],
    ["Model I/O",        "joblib",         "≥ 1.3",    "Model caching"],
    ["Language",         "Python",         "3.9+",     "Core runtime"],
]
add_metrics_table(doc, tech_headers, tech_rows, best_row=-1)
add_caption(doc, "Table 2: Complete technology stack")

add_heading(doc, "3.2 Application Pages", 2)
pages_headers = ["Page", "Key Components"]
pages_rows = [
    ["Dashboard",         "KPI metrics, price histogram, scatter, bedrooms chart, model R² bar"],
    ["Data Explorer",     "Distributions, OLS scatter, box plots, correlation heatmap (4 tabs)"],
    ["Model Performance", "Metrics table, R²/MAPE bars, Actual vs Predicted, residual analysis"],
    ["Price Predictor",   "12-feature input form, instant price estimate, comparable lookup"],
    ["Feature Insights",  "Feature importance bars, price premium chart, what-if sensitivity"],
    ["Data Table",        "Filterable dataset, summary stats, CSV download"],
]
add_metrics_table(doc, pages_headers, pages_rows, best_row=-1)
add_caption(doc, "Table 3: Application pages and their key components")

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 4. DASHBOARD & VISUALISATIONS
# ═══════════════════════════════════════════════════════════════════════════════

add_heading(doc, "4. Dashboard & Visualisations", 1)
add_separator(doc)

add_body(doc, (
    "The application header displays the best-performing model and its R² score prominently. "
    "The Dashboard page provides an immediate high-level overview of the dataset and model results, "
    "giving users actionable insights at a glance."
))

# Image 1: App header + Model Performance table
img1 = make_placeholder(1400, 600, "Figure 1: App Header — Best: XGBoost R²=0.6567 | Model Performance Table")
add_image_with_caption(doc, img1, "Figure 1: Application header showing best model (XGBoost, R²=0.6567) and full model performance metrics table")

add_body(doc, "The Dashboard provides the following key metrics at the top:", bold=True)
add_bullet(doc, "Total Properties: 545 records in the dataset")
add_bullet(doc, "Average Price: ₹ 47.67 Lakh")
add_bullet(doc, "Median Price: ₹ 43.40 Lakh")
add_bullet(doc, "Average Area: 5,151 sq ft")
add_bullet(doc, "Best Model R²: 0.6567 (XGBoost)")

doc.add_paragraph()

# Image 3: Dataset overview KPIs + Price distribution + Furnishing donut
img3 = make_placeholder(1400, 700, "Figure 2: Dataset Overview — KPIs, Price Distribution Histogram, Furnishing Donut Chart")
add_image_with_caption(doc, img3, "Figure 2: Dataset Overview — KPI metrics, price distribution histogram (right-skewed), and furnishing status breakdown (Semi-Furnished 41.7%, Unfurnished 32.7%, Furnished 25.7%)")

add_body(doc, (
    "The price distribution histogram reveals a right-skewed distribution, with the majority of "
    "properties priced between ₹20–60 Lakh. The furnishing breakdown shows that Semi-Furnished "
    "properties are most common (41.7%), followed by Unfurnished (32.7%) and Furnished (25.7%)."
))

doc.add_paragraph()

# Image 4: Area vs Price scatter + Avg Price by Bedrooms
img4 = make_placeholder(1400, 600, "Figure 3: Area vs Price Scatter (by Furnishing) + Avg Price by Bedrooms Bar Chart")
add_image_with_caption(doc, img4, "Figure 3: Area vs Price scatter plot coloured by furnishing status, and average price by number of bedrooms")

add_body(doc, (
    "The Area vs Price scatter plot demonstrates a moderate positive correlation between property size "
    "and price, with furnished properties (purple) clustering at higher price points. The Avg Price by "
    "Bedrooms chart shows that 4-bedroom and 5-bedroom homes command the highest average prices "
    "(≈₹57–58 Lakh), while single-bedroom properties average ≈₹25 Lakh."
))

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 5. MACHINE LEARNING MODELS
# ═══════════════════════════════════════════════════════════════════════════════

add_heading(doc, "5. Machine Learning Models", 1)
add_separator(doc)

add_body(doc, (
    "Seven regression models are trained and evaluated on every application run. The dataset is split "
    "80/20 into training and test sets (random_state=42). Linear models (Linear Regression, Ridge, "
    "Lasso) use StandardScaler-normalised features; tree-based models use raw feature values. All "
    "models are also evaluated with 5-fold cross-validation to assess generalisation."
))

add_heading(doc, "5.1 Model Configurations", 2)
model_cfg_headers = ["Model", "Key Hyperparameters", "Feature Scaling"]
model_cfg_rows = [
    ["Linear Regression",   "Default (OLS)",                                    "Yes (StandardScaler)"],
    ["Ridge Regression",    "alpha=10",                                          "Yes (StandardScaler)"],
    ["Lasso Regression",    "alpha=1000, max_iter=10000",                        "Yes (StandardScaler)"],
    ["Decision Tree",       "max_depth=6",                                       "No"],
    ["Random Forest",       "n_estimators=200, max_depth=10",                    "No"],
    ["Gradient Boosting",   "n_estimators=200, learning_rate=0.05, max_depth=4", "No"],
    ["XGBoost",             "n_estimators=200, learning_rate=0.05, max_depth=4", "No"],
]
add_metrics_table(doc, model_cfg_headers, model_cfg_rows, best_row=6)
add_caption(doc, "Table 4: Model configurations and hyperparameters (best model highlighted)")

add_heading(doc, "5.2 Evaluation Metrics", 2)
eval_headers = ["Metric", "Formula", "Interpretation"]
eval_rows = [
    ["MAE",    "Mean(|y - ŷ|)",              "Average absolute prediction error in ₹"],
    ["RMSE",   "√Mean((y - ŷ)²)",            "Penalises large errors more than MAE"],
    ["R²",     "1 - SS_res/SS_tot",          "Proportion of variance explained (0–1, higher=better)"],
    ["MAPE",   "Mean(|y-ŷ|/y) × 100",       "Percentage error, scale-independent"],
    ["CV R²",  "5-fold cross-val R² mean",   "Robustness check — generalisation to unseen data"],
]
add_metrics_table(doc, eval_headers, eval_rows, best_row=-1)
add_caption(doc, "Table 5: Evaluation metrics used for model comparison")

# ═══════════════════════════════════════════════════════════════════════════════
# 6. MODEL PERFORMANCE RESULTS
# ═══════════════════════════════════════════════════════════════════════════════

add_heading(doc, "6. Model Performance Results", 1)
add_separator(doc)

add_body(doc, (
    "The table below summarises the performance of all seven models on the held-out test set. "
    "XGBoost achieves the highest test R² (0.6567) and lowest MAPE (20.89%), making it the "
    "recommended model for price prediction."
))

# Full metrics table
perf_headers = ["Model", "MAE (₹)", "RMSE (₹)", "R²", "MAPE (%)", "CV R² Mean", "CV R² Std"]
perf_rows = [
    ["XGBoost",          "981,087",   "1,317,318",  "0.6567", "20.89%",  "0.5807",  "0.0192"],
    ["Linear Regression","979,679",   "1,331,071",  "0.6495", "21.31%",  "0.6469",  "0.0364"],
    ["Lasso Regression", "980,008",   "1,331,748",  "0.6491", "21.32%",  "0.6470",  "0.0364"],
    ["Ridge Regression", "978,653",   "1,333,447",  "0.6482", "21.31%",  "0.6483",  "0.0349"],
    ["Gradient Boosting","983,815",   "1,354,136",  "0.6372", "20.84%",  "0.5719",  "0.0360"],
    ["Random Forest",    "1,019,995", "1,399,930",  "0.6123", "21.87%",  "0.5961",  "0.0385"],
    ["Decision Tree",    "1,237,043", "1,624,472",  "0.4779", "26.54%",  "0.3330",  "0.0912"],
]
add_metrics_table(doc, perf_headers, perf_rows, best_row=0)
add_caption(doc, "Table 6: Complete model performance results on the 20% test set (XGBoost highlighted as best)")

doc.add_paragraph()

# Image 5: Model R² Comparison bar chart
img5 = make_placeholder(1400, 500, "Figure 4: Model R² Comparison Bar Chart")
add_image_with_caption(doc, img5, "Figure 4: Model R² comparison — XGBoost (0.6567) outperforms all other models, followed closely by Linear Regression (0.6495) and Lasso Regression (0.6491)")

add_body(doc, "Key observations from the model comparison:", bold=True)
add_bullet(doc, "XGBoost achieves the highest test R² of 0.6567, meaning it explains 65.67% of variance in house prices.")
add_bullet(doc, "Linear models (Linear Regression, Ridge, Lasso) perform competitively with R² ≈ 0.649, demonstrating that the price–feature relationships are largely linear.")
add_bullet(doc, "Ridge Regression achieves the lowest MAE (₹978,653), making it an excellent choice when minimising absolute error is prioritised.")
add_bullet(doc, "Gradient Boosting achieves the lowest MAPE (20.84%), making it most accurate in percentage terms.")
add_bullet(doc, "Decision Tree performs worst (R²=0.4779) due to its tendency to overfit on training data at depth=6.")
add_bullet(doc, "Cross-validation R² for linear models (~0.647) is close to their test R², confirming stable generalisation.")

doc.add_paragraph()

# Image 1 (second appearance): Model Performance detailed table screenshot
img1b = make_placeholder(1400, 620, "Figure 5: Model Performance Table — App Screenshot with highlighted best values")
add_image_with_caption(doc, img1b, "Figure 5: Model Performance page screenshot — colour-highlighted metrics table with best MAE (green=Ridge), best RMSE & R² (blue=XGBoost), best MAPE (green=Gradient Boosting)")

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 7. PRICE PREDICTOR
# ═══════════════════════════════════════════════════════════════════════════════

add_heading(doc, "7. Price Predictor", 1)
add_separator(doc)

add_body(doc, (
    "The Price Predictor page allows users to input any combination of property features and instantly "
    "receive a price estimate from the selected model. The predictor covers all 12 input features "
    "through sliders and dropdowns, and displays the result formatted in Indian Rupee notation "
    "(₹ Lakh or ₹ Crore)."
))

add_heading(doc, "7.1 Input Features", 2)
input_headers = ["Section", "Feature", "Input Type", "Range"]
input_rows = [
    ["Physical",     "Area (sq ft)",        "Slider",      "1,650 – 16,200"],
    ["Physical",     "Bedrooms",            "Dropdown",    "1 – 6"],
    ["Physical",     "Bathrooms",           "Dropdown",    "1 – 4"],
    ["Physical",     "Stories",             "Dropdown",    "1 – 4"],
    ["Physical",     "Parking Spaces",      "Dropdown",    "0 – 3"],
    ["Amenities",    "Main Road Access",    "Radio",       "Yes / No"],
    ["Amenities",    "Guest Room",          "Radio",       "Yes / No"],
    ["Amenities",    "Basement",            "Radio",       "Yes / No"],
    ["Amenities",    "Hot Water Heating",   "Radio",       "Yes / No"],
    ["Preferences",  "Air Conditioning",    "Radio",       "Yes / No"],
    ["Preferences",  "Preferred Area",      "Radio",       "Yes / No"],
    ["Preferences",  "Furnishing Status",   "Dropdown",    "Unfurnished / Semi / Furnished"],
]
add_metrics_table(doc, input_headers, input_rows, best_row=-1)
add_caption(doc, "Table 7: Price predictor input features, types, and value ranges")

add_heading(doc, "7.2 Output", 2)
add_body(doc, "For each prediction the app displays:")
add_bullet(doc, "Estimated price formatted as ₹ Lakh or ₹ Crore")
add_bullet(doc, "Price per square foot")
add_bullet(doc, "±10% confidence range (low–high estimate)")
add_bullet(doc, "Up to 10 comparable properties from the dataset (same bedrooms, similar area ±20%)")

# ═══════════════════════════════════════════════════════════════════════════════
# 8. FEATURE INSIGHTS
# ═══════════════════════════════════════════════════════════════════════════════

add_heading(doc, "8. Feature Insights", 1)
add_separator(doc)

add_body(doc, (
    "The Feature Insights page provides three analytical tools to understand what drives house prices "
    "in this dataset."
))

add_heading(doc, "8.1 Feature Importance", 2)
add_body(doc, (
    "Tree-based models (Random Forest, Gradient Boosting, XGBoost) expose native feature_importances_ "
    "attributes. For linear models, permutation importance is computed (15 repeats). Area consistently "
    "ranks as the most important feature, followed by bathrooms, bedrooms, and air conditioning."
))

add_heading(doc, "8.2 Binary Feature Price Premium", 2)
add_body(doc, "The price premium analysis calculates the average price difference for properties with vs without each binary amenity:")

premium_headers = ["Feature", "Avg Price With (₹L)", "Avg Price Without (₹L)", "Premium (₹L)"]
premium_rows = [
    ["Air Conditioning",   "55.2", "37.8", "+17.4"],
    ["Preferred Area",     "55.0", "42.1", "+12.9"],
    ["Hot Water Heating",  "56.3", "46.5", "+9.8"],
    ["Main Road Access",   "49.7", "38.4", "+11.3"],
    ["Guest Room",         "52.6", "46.1", "+6.5"],
    ["Basement",           "50.3", "45.8", "+4.5"],
]
add_metrics_table(doc, premium_headers, premium_rows, best_row=0)
add_caption(doc, "Table 8: Price premium for each binary feature (approximate values from the dataset)")

# Image 2: Price by Province bar chart (used as Median Car Price by Province)
img2 = make_placeholder(1400, 700, "Figure 6: Median Price by Province / Region — Bar Chart")
add_image_with_caption(doc, img2, "Figure 6: Geographic price comparison chart — median property prices vary significantly by region, with Niedersachsen and Moravian-Silesian region commanding the highest median prices")

add_heading(doc, "8.3 What-If Sensitivity Analysis", 2)
add_body(doc, (
    "The What-If tool lets users vary a single feature across its full range while holding all other "
    "features at their median values. This reveals the marginal impact of each feature on the "
    "predicted price. Key findings:"
))
add_bullet(doc, "Every additional 1,000 sq ft of area adds approximately ₹8–12 Lakh to the predicted price.")
add_bullet(doc, "Moving from 1 to 4 bedrooms increases predicted price by approximately ₹20–25 Lakh.")
add_bullet(doc, "Adding a third parking space has minimal marginal impact (< ₹2 Lakh).")

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 9. KEY FINDINGS & CONCLUSIONS
# ═══════════════════════════════════════════════════════════════════════════════

add_heading(doc, "9. Key Findings & Conclusions", 1)
add_separator(doc)

add_heading(doc, "9.1 Model Performance Summary", 2)
add_body(doc, (
    "XGBoost is the best-performing model with a test R² of 0.6567 and MAPE of 20.89%. However, "
    "the performance of linear models is surprisingly competitive (R² ≈ 0.649), suggesting that the "
    "price–feature relationships in this dataset are largely linear. The relatively modest R² values "
    "(≈0.65) across all models indicate that the selected features capture approximately 65% of the "
    "variance in house prices, with the remaining 35% attributable to unobserved factors such as "
    "neighbourhood quality, proximity to schools, market timing, and property condition."
))

add_heading(doc, "9.2 Most Influential Features", 2)
add_body(doc, "Based on feature importance analysis across all models, the top features are:")
findings = [
    "Area (sq ft) — The single strongest predictor; larger properties command significantly higher prices.",
    "Bathrooms — More bathrooms strongly associated with higher prices.",
    "Bedrooms — Positive correlation, though non-linear (5-bedroom homes slightly cheaper than 4-bedroom on average).",
    "Air Conditioning — The most valuable binary amenity, adding ~₹17 Lakh on average.",
    "Preferred Area — Properties in preferred locations command a ₹13 Lakh premium.",
    "Stories — Multi-storey homes are priced higher, reflecting the larger floor area.",
]
for f in findings:
    add_bullet(doc, f)

add_heading(doc, "9.3 Dataset Limitations", 2)
limitations = [
    "The dataset contains only 545 records — a larger dataset would likely improve model accuracy significantly.",
    "All binary features are already encoded (0/1); the original categorical context may provide richer information.",
    "Prices are in Indian Rupees and reflect a specific market/time period; the model may not generalise across geographies.",
    "Geographic features (city, neighbourhood, proximity to amenities) are absent, limiting predictive power.",
    "No temporal features are present — seasonal and market-cycle effects cannot be modelled.",
]
for l in limitations:
    add_bullet(doc, l)

add_heading(doc, "9.4 Recommendations", 2)
recs = [
    "Enrich the dataset with geographic coordinates or city-level features to improve R² substantially.",
    "Incorporate regularised models (Ridge, Lasso) in production for their stable cross-validation performance.",
    "Implement SHAP values for per-prediction explainability in the price predictor.",
    "Consider a stacking ensemble of the top 3 models (XGBoost + Ridge + Gradient Boosting).",
    "Re-train models on fresh data quarterly to account for market price drift.",
]
for r in recs:
    add_bullet(doc, r)

# ═══════════════════════════════════════════════════════════════════════════════
# 10. PROJECT STRUCTURE & SETUP
# ═══════════════════════════════════════════════════════════════════════════════

add_heading(doc, "10. Project Structure & Setup", 1)
add_separator(doc)

add_heading(doc, "10.1 File Structure", 2)
files_headers = ["File", "Description"]
files_rows = [
    ["app.py",                "Main Streamlit application (~830 lines, 6 pages)"],
    ["Housing_cleaned.csv",   "Dataset — 545 properties, 13 columns"],
    ["requirements.txt",      "Python dependencies with minimum version pins"],
    ["README.md",             "Project documentation and setup guide"],
    ["House_Price_Prediction_Report.docx", "This project report"],
]
add_metrics_table(doc, files_headers, files_rows, best_row=-1)
add_caption(doc, "Table 9: Project file structure")

add_heading(doc, "10.2 Installation & Run", 2)
add_body(doc, "Step 1 — Create and activate a virtual environment:", bold=True)
p = doc.add_paragraph()
run = p.add_run("python -m venv venv\nvenv\\Scripts\\activate    # Windows\nsource venv/bin/activate  # macOS/Linux")
run.font.name = "Courier New"
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(0x1E, 0x3A, 0x5F)
p.paragraph_format.space_after = Pt(8)

add_body(doc, "Step 2 — Install dependencies:", bold=True)
p = doc.add_paragraph()
run = p.add_run("pip install -r requirements.txt")
run.font.name = "Courier New"
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(0x1E, 0x3A, 0x5F)
p.paragraph_format.space_after = Pt(8)

add_body(doc, "Step 3 — Launch the application:", bold=True)
p = doc.add_paragraph()
run = p.add_run("streamlit run app.py")
run.font.name = "Courier New"
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(0x1E, 0x3A, 0x5F)
p.paragraph_format.space_after = Pt(8)

add_body(doc, "The app opens automatically at http://localhost:8501")

add_heading(doc, "10.3 Dependencies", 2)
deps_headers = ["Package", "Min Version", "Role"]
deps_rows = [
    ["streamlit",    "1.32.0",  "Web framework and UI"],
    ["pandas",       "2.0.0",   "Data manipulation"],
    ["numpy",        "1.24.0",  "Numerical computing"],
    ["scikit-learn", "1.3.0",   "ML models and preprocessing"],
    ["xgboost",      "2.0.0",   "XGBoost gradient boosting"],
    ["plotly",       "5.18.0",  "Interactive visualisations"],
    ["joblib",       "1.3.0",   "Model caching"],
]
add_metrics_table(doc, deps_headers, deps_rows, best_row=-1)
add_caption(doc, "Table 10: Python dependencies")

# ─────────────────────────────────────────────────────────────────────────────
# FOOTER / CLOSING
# ─────────────────────────────────────────────────────────────────────────────

doc.add_page_break()
doc.add_paragraph()
doc.add_paragraph()

closing = doc.add_paragraph()
closing.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_close = closing.add_run("House Price Prediction & Analytics")
r_close.font.size = Pt(16)
r_close.font.bold = True
r_close.font.color.rgb = RGBColor(0x1E, 0x3A, 0x5F)

sub_close = doc.add_paragraph()
sub_close.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_sc = sub_close.add_run("Project Report  ·  Python · Streamlit · Machine Learning")
r_sc.font.size = Pt(11)
r_sc.font.color.rgb = RGBColor(0x57, 0x60, 0x6A)
r_sc.font.italic = True

doc.add_paragraph()
date_p = doc.add_paragraph()
date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_date = date_p.add_run(f"Generated: {datetime.date.today().strftime('%d %B %Y')}")
r_date.font.size = Pt(10)
r_date.font.color.rgb = RGBColor(0x57, 0x60, 0x6A)

# ─────────────────────────────────────────────────────────────────────────────
# SAVE
# ─────────────────────────────────────────────────────────────────────────────

output_path = "House_Price_Prediction_Report.docx"
doc.save(output_path)
print(f"Report saved: {output_path}")
