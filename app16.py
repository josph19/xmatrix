import streamlit as st
import requests
import pandas as pd
import re
from io import BytesIO

# GroqCloud API key and model
API_KEY = "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"
MODEL = "llama-3.1-8b-instant"

# Sidebar mode selection
mode = st.sidebar.radio("Matrix Generation Mode", ["Automatic (AI)", "Manual"])

# Function to send request to GroqCloud
def query_groq(prompt):
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    data = {
        "model": MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.7
    }
    response = requests.post("https://api.groq.com/openai/v1/chat/completions", headers=headers, json=data)
    if response.status_code == 200:
        return response.json()["choices"][0]["message"]["content"]
    else:
        raise Exception(f"Error: {response.status_code} - {response.text}")

# Function to parse markdown tables into DataFrames
def parse_markdown_tables(markdown_text):
    tables = []
    table_blocks = re.findall(r'((?:\|[^\n]*\|\n)+)', markdown_text)
    for block in table_blocks:
        lines = block.strip().split('\n')
        if not lines:
            continue
        rows = []
        for line in lines:
            if line.startswith('|') and line.endswith('|'):
                if re.match(r'\|\s*[-:]+\s*\|', line):
                    continue
                cells = [cell.strip() for cell in line.split('|')[1:-1]]
                rows.append(cells)
        if rows and all(len(row) == len(rows[0]) for row in rows):
            if len(rows) > 1:
                headers = rows[0]
                data = rows[1:]
                df = pd.DataFrame(data, columns=headers)
                tables.append(df)
            else:
                df = pd.DataFrame([rows[0]])
                tables.append(df)
    return tables

# ---------------------------------------------
# AUTOMATIC MODE
# ---------------------------------------------
if mode == "Automatic (AI)":
    st.title("Hoshin Kanri Matrix Generator")

    strategic_objective = st.text_area("Strategic Objective(s)", "")
    annual_objectives = st.text_area("Annual Objectives", "")
    improvement_priorities = st.text_area("Improvement Priorities", "")
    kpis = st.text_area("Key Performance Indicators (KPIs)", "")
    responsibilities = st.text_area("Responsibilities", "")

    if st.button("Generate X Matrix"):
        if all([strategic_objective, annual_objectives, improvement_priorities, kpis, responsibilities]):
            with st.spinner("Generating matrix..."):
                prompt_matrix = f"""
You are an expert in strategic planning. Create a Hoshin Kanri matrix in the form of several tables. Here are the inputs:

Strategic Objective:
{strategic_objective}

Annual Objectives:
{annual_objectives}

Improvement Priorities:
{improvement_priorities}

KPIs:
{kpis}

Responsibilities:
{responsibilities}

Expected table structure:

1. **Table 1**:
   - Rows: Strategic Objectives
   - Columns: Annual Objectives
   - Cells: 'O' = primary relation, 'X' = secondary relation, empty = no relation

2. **Table 2**:
   - Rows: Annual Objectives
   - Columns: Improvement Priorities
   - Cells: 'O', 'X' or empty

3. **Table 3**:
   - Rows: Improvement Priorities
   - Columns: KPIs
   - Cells: 'O', 'X' or empty

4. **Table 4**:
   - Rows: Improvement Priorities
   - Columns: Responsible Persons

Generate these tables with clear headers and good structure.
"""
                try:
                    result_matrix = query_groq(prompt_matrix)
                    st.markdown("### Generated Hoshin Kanri Matrix:")
                    st.markdown(result_matrix)
                    st.session_state.result_matrix = result_matrix

                    st.markdown("""
> **Legend**:
>
> Empty cells mean no relation.  
> **'O'** means a **primary relation**, **'X'** means a **secondary relation**.
""")

                    prompt_suggestions = f"""
Here are the defined elements in the strategic plan:

- Strategic Objective: {strategic_objective}
- Annual Objectives: {annual_objectives}
- Improvement Priorities: {improvement_priorities}
- KPIs: {kpis}

Give concrete and well-argued improvement suggestions for each of the following:
1. Annual Objectives: relevance, clarity, strategic alignment.
2. Improvement Priorities: operational impact, feasibility.
3. KPIs: accuracy, relevance, usefulness for monitoring.

Do not start each point with generic intros. Provide focused bullet-point recommendations.
"""
                    result_suggestions = query_groq(prompt_suggestions)
                    st.markdown("---")
                    st.markdown("### Improvement Suggestions:")
                    st.markdown(result_suggestions)
                    st.session_state.result_suggestions = result_suggestions

                    st.markdown("---")
                    st.markdown("### Download Report")
                    tables = parse_markdown_tables(result_matrix)

                    if not tables or len(tables) < 4:
                        st.warning("Unable to generate Excel file. Incorrect format.")
                    else:
                        buffer = BytesIO()
                        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                            table_names = [
                                "Strategic-Annual",
                                "Annual-Priorities",
                                "KPIs-Priorities",
                                "Priorities-Responsibilities"
                            ]
                            for i, table in enumerate(tables[:4]):
                                table.to_excel(writer, sheet_name=table_names[i], index=False)

                            try:
                                obj_annual_match = re.search(r'1\.\s*Annual Objectives[^\n]*\n((?:.+\n)+?)(?:\n*2\.|\Z)',
                                                             result_suggestions, re.DOTALL)
                                obj_annual = obj_annual_match.group(1).strip() if obj_annual_match else "Not available"

                                priorities_match = re.search(r'2\.\s*Improvement Priorities[^\n]*\n((?:.+\n)+?)(?:\n*3\.|\Z)',
                                                             result_suggestions, re.DOTALL)
                                priorities = priorities_match.group(1).strip() if priorities_match else "Not available"

                                kpis_match = re.search(r'3\.\s*KPIs[^\n]*\n((?:.+\n)+?)(?:\Z)',
                                                       result_suggestions, re.DOTALL)
                                kpi_section = kpis_match.group(1).strip() if kpis_match else "Not available"

                                suggestions_df = pd.DataFrame({
                                    "Category": ["Annual Objectives", "Improvement Priorities", "KPIs"],
                                    "Suggestions": [obj_annual, priorities, kpi_section]
                                })
                            except Exception:
                                suggestions_df = pd.DataFrame({
                                    "Suggestions": ["Error extracting suggestions."]
                                })

                            suggestions_df.to_excel(writer, sheet_name="Suggestions", index=False)
                        buffer.seek(0)

                        st.download_button(
                            label="Download Excel Report",
                            data=buffer,
                            file_name="Hoshin_Kanri_Matrix.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="excel_download"
                        )
                except Exception as e:
                    st.error(f"Error during generation: {e}")
        else:
            st.warning("Please fill in all fields to generate the matrix.")

# ---------------------------------------------
# MANUAL MODE
# ---------------------------------------------
elif mode == "Manual":
    st.title("Manual Hoshin Kanri Matrix Builder")

    st.subheader("Strategic Objectives")
    strategic_goals = st.text_area("List strategic objectives (one per line)").split("\n")
    strategic_goals = [g.strip() for g in strategic_goals if g.strip()]

    st.subheader("Annual Objectives")
    annual_goals = st.text_area("List annual objectives (one per line)").split("\n")
    annual_goals = [g.strip() for g in annual_goals if g.strip()]

    st.subheader("Strategic and Annual Objectives Relations")
    strat_annual_matrix = {}
    for strat_goal in strategic_goals:
        strat_annual_matrix[strat_goal] = {}
        for ann_goal in annual_goals:
            relation = st.selectbox(
                f"Relation between strategic '{strat_goal}' and annual '{ann_goal}'",
                ["", "O", "X"],
                key=f"strat_{strat_goal}_{ann_goal}"
            )
            strat_annual_matrix[strat_goal][ann_goal] = relation.split()[0] if relation else ""

    st.subheader("Improvement Priorities")
    improvement_priorities = st.text_area("List priorities (one per line)").split("\n")
    improvement_priorities = [p.strip() for p in improvement_priorities if p.strip()]

    st.subheader("Annual Objectives and Priorities Relations")
    annual_priority_matrix = {}
    for goal in annual_goals:
        annual_priority_matrix[goal] = {}
        for priority in improvement_priorities:
            rel = st.selectbox(
                f"Relation between annual objective '{goal}' and priority '{priority}'", ["", "O", "X"],
                key=f"ann_{goal}_{priority}"
            )
            annual_priority_matrix[goal][priority] = rel

    st.subheader("KPIs")
    kpis = st.text_area("List KPIs (one per line)").split("\n")
    kpis = [k.strip() for k in kpis if k.strip()]

    st.subheader("Priorities and KPIs Relations")
    priority_kpi_matrix = {}
    for priority in improvement_priorities:
        priority_kpi_matrix[priority] = {}
        for kpi in kpis:
            rel = st.selectbox(
                f"Relation between priority '{priority}' and KPI '{kpi}'", ["", "O", "X"],
                key=f"priority_{priority}_{kpi}"
            )
            priority_kpi_matrix[priority][kpi] = rel

    st.subheader("Responsible Persons")
    responsables = st.text_area("List responsible persons (one per line)").split("\n")
    responsables = [r.strip() for r in responsables if r.strip()]

    st.subheader("Priorities and Responsibilities Relations")
    priority_responsible_matrix = {}
    for priority in improvement_priorities:
        priority_responsible_matrix[priority] = {}
        for resp in responsables:
            rel = st.selectbox(
                f"Responsibility of '{resp}' for '{priority}'", ["", "O", "X"],
                key=f"resp_{priority}_{resp}"
            )
            priority_responsible_matrix[priority][resp] = rel

    st.subheader("Download X Matrix")
    if st.button("Export to Excel"):
        df1 = pd.DataFrame.from_dict(strat_annual_matrix, orient='index')
        df2 = pd.DataFrame.from_dict(annual_priority_matrix, orient='index')
        df3 = pd.DataFrame.from_dict(priority_kpi_matrix, orient='index')
        df4 = pd.DataFrame.from_dict(priority_responsible_matrix, orient='index')

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df1.to_excel(writer, sheet_name="Strategic-Annual")
            df2.to_excel(writer, sheet_name="Annual-Priorities")
            df3.to_excel(writer, sheet_name="Priorities-KPIs")
            df4.to_excel(writer, sheet_name="Priorities-Responsibilities")
        buffer.seek(0)

        st.download_button(
            label="Download Excel File",
            data=buffer,
            file_name="Manual_Hoshin_Kanri_Matrix.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.markdown(
    """
    <hr style="margin-top: 50px;"/>
    <div style="text-align: center; color: gray; font-size: 0.9em;">
        Made by <strong>joseph19</strong>
    </div>
    """,
    unsafe_allow_html=True
)
