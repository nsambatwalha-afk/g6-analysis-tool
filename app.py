import streamlit as st
import numpy as np
import pandas as pd
import libfunc
import truss_analysis
from extras import *

def steel_input_ui(val=False):
    if not val:
        st.subheader("Steel Properties")

        truss_analysis.grade = st.selectbox("Steel Grade", ["S235", "S275", "S355", "S420", "S450", "S460"])

        truss_analysis.shapeten = st.selectbox("Tension Member Shape", ["Angle", "CHS"])

        truss_analysis.shapecomp = st.selectbox("Compression Member Shape", ["UB", "UC", "CHS", "Angle"])

        truss_analysis.jointing = st.selectbox("Joint Type", ["bolt", "weld"])

        if truss_analysis.jointing == "bolt":
            truss_analysis.nh = st.number_input("Number of bolt holes", value=2.0)
            truss_analysis.d = st.number_input("Bolt diameter (mm)", value=20.0)

            truss_analysis.stag = st.checkbox("Staggered bolts?")

            if truss_analysis.stag:
                truss_analysis.s = st.number_input("Stagger spacing s (mm)", value=50.0)
                truss_analysis.p = st.number_input("Pitch p (mm)", value=100.0)
                truss_analysis.ngs = st.number_input("Number of stagger lines", value=1.0)
            else:
                truss_analysis.s, truss_analysis.p, truss_analysis.ngs = 0, 0, 0

        else:
            truss_analysis.nh = truss_analysis.d = truss_analysis.s = truss_analysis.p = truss_analysis.ngs = 0
            truss_analysis.stag = False

        return
    else:
        st.subheader("Steel Properties")
        truss_analysis.grade = st.selectbox("Steel Grade", ["S235", "S275", "S355", "S420", "S450", "S460"])


st.set_page_config(page_title="Structural Engineering Toolkit", layout="wide")

st.title("🏗 Structural Engineering Toolkit")

# ---- Sidebar Navigation ----
task = st.sidebar.radio(
    "Select Task",
    (
        "Matrix Multiplication",
        "Gauss-Jordan Elimination",
        "Truss Analysis & Design",
        "Single Member Design",
        "Simple Beam Design",
        "Beam Analysis & Design"
    )
)

# =====================================================
# MATRIX MULTIPLICATION
# =====================================================

if task == "Matrix Multiplication":

    st.header("Matrix Multiplication")

    m1 = st.text_area("Enter Matrix 1 (example: 1,2;3,4)")
    m2 = st.text_area("Enter Matrix 2 (example: 5,6;7,8)")

    if st.button("Multiply"):

        try:
            arg1 = np.matrix(m1)
            arg2 = np.matrix(m2)

            if libfunc.checker(arg1, arg2):
                result = libfunc.multiplier(arg1, arg2)
                st.success("Result:")
                st.write(result)
            else:
                st.error("Incompatible matrices.")

        except Exception as e:
            st.error(f"Error: {e}")

# =====================================================
# GAUSS-JORDAN ELIMINATION
# =====================================================

elif task == "Gauss-Jordan Elimination":

    st.header("Gauss-Jordan Elimination")

    aug = st.text_area("Enter Augmented Matrix (example: 1,2,3;4,5,6)")

    if st.button("Solve"):

        try:
            arg = np.matrix(aug)
            solution = libfunc.solver(arg)

            st.success("Solution:")
            st.write(solution)

        except Exception as e:
            st.error(f"Error: {e}")

# =====================================================
# TRUSS ANALYSIS & MEMBER DESIGN
# =====================================================

elif task == "Truss Analysis & Design":

    st.header("Truss Analysis & Member Design")

    st.info(
        "Upload the required Excel files before running the analysis."
    )

    # -------------------------
    # Download Templates
    # -------------------------

    st.subheader("Download Input Templates")

    col1, col2 = st.columns(2)

    try:
        with col1:
            with open("joints_template.xlsx", "rb") as f:
                st.download_button(
                    label="Download joints template",
                    data=f,
                    file_name="joints.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except:
        st.warning("joints template not found on server.")

    try:
        with col2:
            with open("members_template.xlsx", "rb") as f:
                st.download_button(
                    label="Download members template",
                    data=f,
                    file_name="members.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except:
        st.warning("members template not found on server.")


    st.write("---")

    # -------------------------
    # Excel Upload Section
    # -------------------------

    joints_file = st.file_uploader(
        "Upload joints.xlsx",
        type=["xlsx"]
    )

    members_file = st.file_uploader(
        "Upload members.xlsx",
        type=["xlsx"]
    )


    st.write("---")
    steel_input_ui()

    # -------------------------
    # Run Analysis Button
    # -------------------------

    if st.button("Run Analysis & Design"):

        if joints_file is None:
            st.warning("Please upload joints.xlsx")
            st.stop()

        if members_file is None:
            st.warning("Please upload members.xlsx")
            st.stop()

        try:

            tension_table, compression_table = truss_analysis.run_analysis_and_design_table(
                joints_file,
                members_file
            )

            # -------------------------
            # Display Results
            # -------------------------

            if tension_table is not None and not tension_table.empty:

                st.subheader("Tension Member Design Summary")
                st.dataframe(tension_table)

            if compression_table is not None and not compression_table.empty:

                st.subheader("Compression Member Design Summary")
                st.dataframe(compression_table)

            if (
                (tension_table is None or tension_table.empty)
                and
                (compression_table is None or compression_table.empty)
            ):
                st.warning("No members required design.")

        except Exception as e:

            st.error(f"Analysis error: {e}")

elif task == "Single Member Design":

    st.header("Single Member Steel Design")

    st.info(
        "Fill in steel data below."
    )

    steel_input_ui()

    st.write("---")


    # -------------------------
    # User Inputs
    # -------------------------

    force = st.number_input(
        "Axial Force (kN)",
        value=100.0
    )

    member_type = st.selectbox(
        "Member Type",
        ["Tension", "Compression"]
    )

    length = st.number_input(
        "Member Length (mm) — used only for compression",
        value=3000.0
    )

    # -------------------------
    # Run Design
    # -------------------------

    if st.button("Design Member"):

        try:

            if member_type == "Tension":

                section = truss_analysis.ten_designer(force * 1000)

                if section is None:
                    st.warning("No suitable section found.")
                else:

                    df = pd.DataFrame([{
                        "Type": "Tension",
                        "Shape": truss_analysis.shapeten,
                        "Size": section[1],
                        "Thickness": section[2],
                        "Area": section[0]
                    }])

                    st.success("Design Result")
                    st.dataframe(df)

            else:

                section = truss_analysis.comp_designer(force * 1000, length)

                if section is None:
                    st.warning("No suitable section found.")
                else:

                    df = pd.DataFrame([{
                        "Type": "Compression",
                        "Shape": truss_analysis.shapecomp,
                        "Size": section["size"],
                        "Axis": section["axis"],
                        "Capacity Utilization (%)": round((force / (section["NbRd"] / 1000)) * 100, 3),
                        "χ": round(section["chi"], 3),
                        "NbRd (kN)": round(section["NbRd"] / 1000, 3)
                    }])

                    st.success("Design Result")
                    st.dataframe(df)

        except Exception as e:

            st.error(f"Design error: {e}")

elif task == "Simple Beam Design":

    st.header("Beam Design")

    steel_input_ui(True)

    st.write("---")

    M = st.number_input("Bending Moment (kNm)", value=100.0)
    V = st.number_input("Shear Force (kN)", value=50.0)
    L = st.number_input("Beam Length (mm)", value=1500.0)

    beam_type = st.selectbox(
        "Beam Type",
        ["Restrained", "Unrestrained"]
    )
    if beam_type == "Unrestrained":
        truss_analysis.condition = st.selectbox(
            "Beam Condition",
            ["Rolled", "Welded"]
        )
        truss_analysis.endcondition = st.selectbox(
            "Restraint",
            ["Free","Partial","Full","Cantilever"]
        )

    if st.button("Design Beam"):

        try:
            if beam_type == "Restrained":
                result = truss_analysis.restrained_beam(M, V)

            else:
                result = truss_analysis.unrestrained_beam(M, V, L)

            st.success("Design Result")
            st.dataframe(pd.DataFrame([result]))

        except Exception as e:
            st.error(f"Error: {e}")

elif task == "Beam Analysis & Design":

    st.header("Beam Analysis & Design (Moment Distribution)")

    steel_input_ui(True)

    st.write("---")

    # =========================
    # SUPPORT INPUT
    # =========================
    st.subheader("Supports")

    n_supports = st.number_input("Number of Supports", min_value=2, value=2)

    supports = []

    for i in range(int(n_supports)):
        col1, col2 = st.columns(2)

        with col1:
            x = st.number_input(f"Support {i+1} Position (m)", key=f"sx_md{i}")

        with col2:
            typ = st.selectbox(
                f"Support {i+1} Type",
                ["Pinned", "Fixed"],
                key=f"st_md{i}"
            )

        supports.append((x, typ.lower()))

    supports = sorted(supports, key=lambda x: x[0])

    # =========================
    # SPANS
    # =========================
    spans = []
    for i in range(len(supports)-1):
        spans.append(supports[i+1][0] - supports[i][0])

    # =========================
    # LOAD INPUT
    # =========================
    st.subheader("Loads")

    n_loads = st.number_input("Number of Loads", min_value=1, value=1)

    loads = []

    for i in range(int(n_loads)):

        ltype = st.selectbox(
            f"Load {i+1} Type",
            ["Point Load", "UDL"],
            key=f"lt_md{i}"
        )

        if ltype == "Point Load":
            P = st.number_input(f"P{i+1} (kN)", key=f"P_md{i}")
            x = st.number_input(f"Position (m)", key=f"Px_md{i}")

            loads.append(("point", P, x))

        else:
            w = st.number_input(f"w{i+1} (kN/m)", key=f"w_md{i}")

            start_span = st.number_input(
                f"Start Span No (1-based)",
                min_value=1,
                max_value=len(spans),
                key=f"ws_md{i}"
            )

            end_span = st.number_input(
                f"End Span No (1-based)",
                min_value=1,
                max_value=len(spans),
                key=f"we_md{i}"
            )

            loads.append(("udl", w, int(start_span), int(end_span)))

    st.write("---")

    # =========================
    # ANALYSIS + DESIGN
    # =========================
    if st.button("Analyze & Design Beam"):

        try:
            import numpy as np
            import math

            # -------------------------
            # END CONDITIONS
            # -------------------------
            end_conditions = []
            for s in supports:
                if s[1] == "fixed":
                    end_conditions.append("fixed")
                else:
                    end_conditions.append("pinned")

            # -------------------------
            # SECTION (EI)
            # -------------------------
            sections = [1.0 for _ in spans]

            # -------------------------
            # LOAD → w PER SPAN
            # -------------------------
            w_span = [0.0]*len(spans)

            for load in loads:

                if load[0] == "udl":
                    _, w, s1, s2 = load
                    for i in range(s1-1, s2):
                        w_span[i] += w

                elif load[0] == "point":
                    _, P, x = load

                    for i in range(len(spans)):
                        a = sum(spans[:i])
                        b = sum(spans[:i+1])

                        if a <= x <= b:
                            L = spans[i]
                            w_span[i] += (2*P)/L   # equivalent UDL

            # -------------------------
            # USE YOUR FUNCTION
            # -------------------------
            w_avg = np.mean(w_span)

            end_moments = moment_dist(
                w_avg,
                end_conditions,
                spans,
                sections
            )

            # -------------------------
            # MAX M + V
            # -------------------------
            def get_max_M_V(end_moments, spans, w_span):

                Mmax = 0
                Vmax = 0

                for i, (Mab, Mba) in enumerate(end_moments):

                    L = spans[i]
                    w = w_span[i]

                    # shear-based critical point
                    try:
                        x = L/2 - (Mba - Mab)/(2*w*L)
                        xs = [0, L, x]
                    except:
                        xs = [0, L]

                    for xi in xs:
                        if 0 <= xi <= L:
                            M = Mab*(1 - xi/L) + Mba*(xi/L) + w*xi*(L-xi)/2
                            Mmax = max(Mmax, abs(M))

                    V = abs((w*L/2) + (Mab + Mba)/L)
                    Vmax = max(Vmax, V)

                return Mmax, Vmax

            M, V = get_max_M_V(end_moments, spans, w_span)

            st.info(f"Max Moment = {round(M,2)} kNm")
            st.info(f"Max Shear = {round(V,2)} kN")

            # -------------------------
            # AUTO RESTRAINT
            # -------------------------
            types = [s[1] for s in supports]

            if all(t == "fixed" for t in types):
                restraint = "Full"
            elif "fixed" in types:
                restraint = "Partial"
            else:
                restraint = "Free"

            # -------------------------
            # DESIGN
            # -------------------------
            if restraint == "Full":
                result = truss_analysis.restrained_beam(M, V)
            else:
                truss_analysis.condition = "Rolled"
                truss_analysis.endcondition = restraint
                result = truss_analysis.unrestrained_beam(M, V, sum(spans)*1000)

            st.success("Design Result")
            st.dataframe(pd.DataFrame([result]))

        except Exception as e:
            st.error(f"Error: {e}")