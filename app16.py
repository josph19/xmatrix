import streamlit as st
import requests
import pandas as pd
import re
from io import BytesIO
import openpyxl 
#from secrets import GROQ_API_KEY

# GroqCloud API key and model
API_KEY = st.secrets["GROQ_API_KEY"]
MODEL = "llama-3.1-8b-instant"

# Sidebar mode selection
mode = st.sidebar.radio("Mode de génération de la matrice", ["Automatique (IA)", "Manuelle"])

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
# AUTOMATED MODE
# ---------------------------------------------
if mode == "Automatique (IA)":
    st.title("Générateur de Matrice Hoshin Kanri")

    strategic_objective = st.text_area("Objectif(s) stratégique(s)", "")
    annual_objectives = st.text_area("Objectifs annuels", "")
    improvement_priorities = st.text_area("Priorités d'amélioration", "")
    kpis = st.text_area("Indicateurs de performance clés (KPIs)", "")
    responsibilities = st.text_area("Responsabilités", "")

    if st.button("Générer la Matrice X"):
        if all([strategic_objective, annual_objectives, improvement_priorities, kpis, responsibilities]):
            with st.spinner("Génération de la matrice..."):
                prompt_matrix = f"""
Tu es un expert en planification stratégique. Crée une matrice Hoshin Kanri sous forme de plusieurs tableaux. Voici les données d'entrée :

Objectif stratégique :
{strategic_objective}

Objectifs annuels :
{annual_objectives}

Priorités d'amélioration :
{improvement_priorities}

KPIs :
{kpis}

Responsabilités :
{responsibilities}

Structure exacte des tableaux en sortie (tous en français) :

1. **Tableau 1** :
   - Lignes : Objectif(s) stratégique(s)
   - Colonnes : Objectifs annuels
   - Cellules : 'O' = relation primaire, 'X' = relation secondaire, vide = aucune relation

2. **Tableau 2** :
   - Lignes : Objectifs annuels
   - Colonnes : Priorités d'amélioration
   - Cellules : 'O', 'X' ou vide comme ci-dessus

3. **Tableau 3** :
   - Lignes : Priorités d'amélioration
   - Colonnes : KPIs (indicateurs de performance clés)
   - Cellules : 'O', 'X' ou vide

4. **Tableau 4** :
   - Lignes : Priorités d'amélioration
   - Colonnes : Responsables

Génère ces tableaux clairement avec des entêtes lisibles et bien structurés.
"""
                try:
                    result_matrix = query_groq(prompt_matrix)
                    st.markdown("### Matrice Hoshin Kanri Générée :")
                    st.markdown(result_matrix)
                    st.session_state.result_matrix = result_matrix

                    st.markdown("""
> **Légende** :
>
> Les cellules vides indiquent qu'il n'y a aucune relation entre les éléments correspondants.  
> Les cellules contenant **'O'** indiquent une **relation primaire**, tandis que les cellules contenant **'X'** indiquent une **relation secondaire**.
""")

                    prompt_suggestions = f"""
Voici les éléments définis dans la planification stratégique :

- Objectif stratégique : {strategic_objective}
- Objectifs annuels : {annual_objectives}
- Priorités d'amélioration : {improvement_priorities}
- KPIs : {kpis}

Propose des suggestions concrètes (autre suggestions) et argumentées pour améliorer chaque point suivant :
1. Objectifs annuels : pertinence, clarté, alignement stratégique.
2. Priorités d'amélioration : impact opérationnel, faisabilité.
3. KPIs : précision, pertinence, suivi utile.

Ne commence pas chaque point par des phrases comme "Définir des objectifs plus..." ou des introductions génériques. Fournis uniquement des recommandations ciblées, sous forme de listes à puces claires.
"""
                    result_suggestions = query_groq(prompt_suggestions)
                    st.markdown("---")
                    st.markdown("### Suggestions d'amélioration :")
                    st.markdown(result_suggestions)
                    st.session_state.result_suggestions = result_suggestions

                    st.markdown("---")
                    st.markdown("### Télécharger le rapport")
                    tables = parse_markdown_tables(result_matrix)

                    if not tables or len(tables) < 4:
                        st.warning("Impossible de générer le fichier Excel. Format incorrect.")
                    else:
                        buffer = BytesIO()
                        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                            table_names = [
                                "Objectifs Strateg-Annuels",
                                "Objectifs Annuels-Priorites",
                                "KPIs-Priorites",
                                "Priorites-Responsables"
                            ]
                            for i, table in enumerate(tables[:4]):
                                table.to_excel(writer, sheet_name=table_names[i], index=False)

                            try:
                                obj_annual_match = re.search(r'1\.\s*Objectifs annuels[^\n]*\n((?:.+\n)+?)(?:\n*2\.|\Z)',
                                                             result_suggestions, re.DOTALL)
                                obj_annual = obj_annual_match.group(1).strip() if obj_annual_match else "Non disponible"

                                priorities_match = re.search(r'2\.\s*Priorités[^\n]*\n((?:.+\n)+?)(?:\n*3\.|\Z)',
                                                             result_suggestions, re.DOTALL)
                                priorities = priorities_match.group(1).strip() if priorities_match else "Non disponible"

                                kpis_match = re.search(r'3\.\s*KPIs[^\n]*\n((?:.+\n)+?)(?:\Z)',
                                                       result_suggestions, re.DOTALL)
                                kpi_section = kpis_match.group(1).strip() if kpis_match else "Non disponible"

                                suggestions_df = pd.DataFrame({
                                    "Catégorie": ["Objectifs Annuels", "Priorités d'Amélioration", "KPIs"],
                                    "Suggestions": [obj_annual, priorities, kpi_section]
                                })
                            except Exception:
                                suggestions_df = pd.DataFrame({
                                    "Suggestions": ["Erreur dans l'extraction des suggestions."]
                                })

                            suggestions_df.to_excel(writer, sheet_name="Suggestions", index=False)
                        buffer.seek(0)

                        st.download_button(
                            label=" Télécharger le rapport Excel",
                            data=buffer,
                            file_name="Matrice_Hoshin_Kanri.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="excel_download"
                        )
                except Exception as e:
                    st.error(f"Erreur lors de la génération : {e}")
        else:
            st.warning("Veuillez remplir tous les champs pour générer la matrice.")

# ---------------------------------------------
# MANUAL MODE
# ---------------------------------------------
# Replace your MANUAL mode block with this updated version
elif mode == "Manuelle":
    st.title("Création manuelle de la matrice Hoshin Kanri")

    st.subheader("Objectifs stratégiques")
    strategic_goals = st.text_area("Liste des objectifs stratégiques (un par ligne)").split("\n")
    strategic_goals = [g.strip() for g in strategic_goals if g.strip()]

    st.subheader("Objectifs annuels")
    annual_goals = st.text_area("Liste des objectifs annuels (un par ligne)").split("\n")
    annual_goals = [g.strip() for g in annual_goals if g.strip()]

    st.subheader("Relations Objectifs stratégiques <-> Objectifs annuels")
    strat_annual_matrix = {}
    for strat_goal in strategic_goals:
        strat_annual_matrix[strat_goal] = {}
        for ann_goal in annual_goals:
            relation = st.selectbox(
                f"Relation entre stratégique '{strat_goal}' et annuel '{ann_goal}'",
                ["", "O (Primaire)", "X (Secondaire)"],
                key=f"strat_{strat_goal}_{ann_goal}"
            )
            strat_annual_matrix[strat_goal][ann_goal] = relation.split()[0] if relation else ""

    st.subheader("Priorités d'amélioration")
    improvement_priorities = st.text_area("Liste des priorités (une par ligne)").split("\n")
    improvement_priorities = [p.strip() for p in improvement_priorities if p.strip()]

    st.subheader("Relations Objectifs annuels <-> Priorités")
    annual_priority_matrix = {}
    for goal in annual_goals:
        annual_priority_matrix[goal] = {}
        for priority in improvement_priorities:
            rel = st.selectbox(
                f"Relation entre objectif annuel '{goal}' et priorité '{priority}'", ["", "O", "X"],
                key=f"ann_{goal}_{priority}"
            )
            annual_priority_matrix[goal][priority] = rel

    st.subheader("KPIs")
    kpis = st.text_area("Liste des KPIs (un par ligne)").split("\n")
    kpis = [k.strip() for k in kpis if k.strip()]

    st.subheader("Relations Priorités <-> KPIs")
    priority_kpi_matrix = {}
    for priority in improvement_priorities:
        priority_kpi_matrix[priority] = {}
        for kpi in kpis:
            rel = st.selectbox(
                f"Relation entre priorité '{priority}' et KPI '{kpi}'", ["", "O", "X"],
                key=f"priority_{priority}_{kpi}"
            )
            priority_kpi_matrix[priority][kpi] = rel

    st.subheader("Responsables")
    responsables = st.text_area("Liste des responsables (un par ligne)").split("\n")
    responsables = [r.strip() for r in responsables if r.strip()]

    st.subheader("Relations Priorités <-> Responsables")
    priority_responsible_matrix = {}
    for priority in improvement_priorities:
        priority_responsible_matrix[priority] = {}
        for resp in responsables:
            rel = st.selectbox(
                f"Responsabilité de '{resp}' pour '{priority}'", ["", "O", "X"],
                key=f"resp_{priority}_{resp}"
            )
            priority_responsible_matrix[priority][resp] = rel

    st.subheader("Télécharger la matrice X")
    if st.button("Exporter au format Excel"):
        df1 = pd.DataFrame.from_dict(strat_annual_matrix, orient='index')
        df2 = pd.DataFrame.from_dict(annual_priority_matrix, orient='index')
        df3 = pd.DataFrame.from_dict(priority_kpi_matrix, orient='index')
        df4 = pd.DataFrame.from_dict(priority_responsible_matrix, orient='index')

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df1.to_excel(writer, sheet_name="Stratégique-Annuels")
            df2.to_excel(writer, sheet_name="Annuels-Priorités")
            df3.to_excel(writer, sheet_name="Priorités-KPIs")
            df4.to_excel(writer, sheet_name="Priorités-Responsables")
        buffer.seek(0)

        st.download_button(
            label="Télécharger le fichier Excel",
            data=buffer,
            file_name="Matrice_Manuelle_Hoshin_Kanri.xlsx",
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
