import os
import pandas as pd
import pdfplumber
import docx
import re
import streamlit as st

# Define scoring criteria
SCORING_WEIGHTS = {
    "Size (EBIT)": 14.3,
    "Market Growth": 14.3,
    "Mission-Critical Offering": 14.3,
    "Recurring Revenue": 14.3,
    "Stable Margins": 14.3,
    "Customer/Supplier Concentration": 14.3,
    "Capital Intensity": 14.3,
}

def extract_text_from_pdf(pdf_file):
    """Extract text from PDF file."""
    text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text

def extract_text_from_docx(docx_file):
    """Extract text from DOCX file."""
    doc = docx.Document(docx_file)
    return "\n".join([para.text for para in doc.paragraphs])

def extract_key_metrics(text):
    """Extract key financial and business metrics from text."""
    metrics = {
        "EBIT": re.search(r"EBIT[^\d]*(\d+[,.]?\d*)", text),
        "Revenue Growth": re.search(r"growth[^\d]*(\d+[,.]?\d*)%", text),
        "EBIT Margins": re.search(r"EBIT margin[^\d]*(\d+[,.]?\d*)%", text),
        "Capex": re.search(r"Capex[^\d]*(\d+[,.]?\d*)", text),
        "Largest Customer %": re.search(r"largest customer[^\d]*(\d+[,.]?\d*)%", text),
    }
    return {key: float(val.group(1).replace(',', '.')) if val else None for key, val in metrics.items()}

def score_deal(metrics):
    """Assign scores based on extracted metrics."""
    scores = {}
    
    # EBIT Scoring
    scores["Size (EBIT)"] = 5 if metrics["EBIT"] and metrics["EBIT"] > 3 else 4 if metrics["EBIT"] > 2 else 3 if metrics["EBIT"] > 1.5 else 2 if metrics["EBIT"] > 1 else 1
    
    # Revenue Growth Scoring
    scores["Market Growth"] = 5 if metrics["Revenue Growth"] and metrics["Revenue Growth"] > 8 else 4 if metrics["Revenue Growth"] > 6 else 3 if metrics["Revenue Growth"] > 5 else 2 if metrics["Revenue Growth"] > 3 else 1
    
    # EBIT Margins Scoring
    scores["Stable Margins"] = 5 if metrics["EBIT Margins"] and metrics["EBIT Margins"] > 20 else 4 if metrics["EBIT Margins"] > 15 else 3 if metrics["EBIT Margins"] > 10 else 2 if metrics["EBIT Margins"] > 5 else 1
    
    # Capex Intensity Scoring
    capex_to_ebitda = (metrics["Capex"] / metrics["EBIT"]) * 100 if metrics["Capex"] and metrics["EBIT"] else None
    scores["Capital Intensity"] = 5 if capex_to_ebitda and capex_to_ebitda < 10 else 4 if capex_to_ebitda < 20 else 3 if capex_to_ebitda < 30 else 2 if capex_to_ebitda < 40 else 1
    
    # Customer Concentration Scoring
    scores["Customer/Supplier Concentration"] = 5 if metrics["Largest Customer %"] and metrics["Largest Customer %"] < 5 else 4 if metrics["Largest Customer %"] < 7 else 3 if metrics["Largest Customer %"] < 10 else 2 if metrics["Largest Customer %"] < 15 else 1
    
    return scores

def calculate_final_score(scores):
    """Calculate weighted final score."""
    weighted_score = sum(scores[criteria] * (SCORING_WEIGHTS[criteria] / 100) for criteria in scores)
    return round(weighted_score, 2)

def main():
    st.title("Deal Evaluation Tool")
    uploaded_file = st.file_uploader("Upload Information Memorandum (PDF/DOCX)", type=["pdf", "docx"])
    
    if uploaded_file:
        with st.spinner("Processing..."):
            text = extract_text_from_pdf(uploaded_file) if uploaded_file.name.endswith("pdf") else extract_text_from_docx(uploaded_file)
            metrics = extract_key_metrics(text)
            scores = score_deal(metrics)
            final_score = calculate_final_score(scores)
        
        st.subheader("Deal Scoring Results")
        st.write(f"Final Score: {final_score} / 5")
        
        df = pd.DataFrame.from_dict(scores, orient='index', columns=['Score'])
        st.table(df)
        
        if final_score >= 4.5:
            st.success("✅ Excellent Deal - Strongly Consider")
        elif final_score >= 4.0:
            st.warning("⚠️ Attractive Deal - Worth Further Analysis")
        elif final_score >= 3.5:
            st.info("ℹ️ Moderate Deal - Needs Deeper Due Diligence")
        else:
            st.error("❌ Weak Deal - Likely Not Worth Pursuing")

if __name__ == "__main__":
    main()
