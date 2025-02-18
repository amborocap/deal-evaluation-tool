import os
import pandas as pd
import pdfplumber
import docx
import re
import streamlit as st

# Define scoring criteria
SCORING_WEIGHTS = {
    "Size (EBIT)": 20.0,
    "Market Growth": 20.0,
    "Mission-Critical Offering": 20.0,
    "Recurring Revenue": 20.0,
    "Stable Margins": 20.0,
}

def extract_text_from_pdf(pdf_file):
    """Extract text from PDF file."""
    text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
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
    }
    
    # Convert extracted values, set defaults if missing
    return {key: float(val.group(1).replace(',', '.')) if val else 0.0 for key, val in metrics.items()}

def score_deal(metrics):
    """Assign scores based on extracted metrics, handling missing values."""
    scores = {}
    
    # EBIT Scoring
    scores["Size (EBIT)"] = (
        5 if metrics["EBIT"] > 3 else
        4 if metrics["EBIT"] > 2 else
        3 if metrics["EBIT"] > 1.5 else
        2 if metrics["EBIT"] > 1 else 1
    )
    
    # Revenue Growth Scoring
    scores["Market Growth"] = (
        5 if metrics["Revenue Growth"] > 8 else
        4 if metrics["Revenue Growth"] > 6 else
        3 if metrics["Revenue Growth"] > 5 else
        2 if metrics["Revenue Growth"] > 3 else 1
    )
    
    # EBIT Margins Scoring
    scores["Stable Margins"] = (
        5 if metrics["EBIT Margins"] > 20 else
        4 if metrics["EBIT Margins"] > 15 else
        3 if metrics["EBIT Margins"] > 10 else
        2 if metrics["EBIT Margins"] > 5 else 1
    )
    
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
