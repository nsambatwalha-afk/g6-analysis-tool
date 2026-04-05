import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import libfunc
import truss_analysis
import report_generator
import section_visualizer
import frame_analysis as fa
from indeterminatebeam import *

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


def _show_section_figure(fig, header="📐 Section Visualization"):
    """Display a section cross-section figure produced by section_visualizer."""
    if fig is not None:
        st.subheader(header)
        st.pyplot(fig, use_container_width=False)
        plt.close(fig)


st.set_page_config(page_title="Structural Engineering Toolkit", layout="wide")

st.title("🏗 Structural Engineering Toolkit")

# ---- Sidebar Navigation ----
task = st.sidebar.radio(
    "Select Task",
    (
        "🏠 Welcome",
        "Matrix Multiplication",
        "Gauss-Jordan Elimination",
        "Truss Analysis & Design",
        "Single Truss Member Design",
        "Simple Beam Design",
        "Beam Analysis & Design",
        "Beam-Column Design",
        "Frame Analysis & Design",
    )
)

# =====================================================
# WELCOME PAGE
# =====================================================

if task == "🏠 Welcome":

    # ── Hero banner ───────────────────────────────────────────────────────────
    st.markdown(
        """
        <div style="
            background: linear-gradient(135deg, #1a3a5c 0%, #2e6da4 100%);
            border-radius: 12px;
            padding: 2.5rem 2rem 2rem 2rem;
            margin-bottom: 1.5rem;
            color: white;
            text-align: center;
        ">
            <h1 style="margin:0; font-size:2.6rem; letter-spacing:1px;">
                🏗 Structural Engineering Toolkit
            </h1>
            <p style="margin:0.5rem 0 0 0; font-size:1.15rem; opacity:0.88;">
                A web-based steel structural analysis &amp; design tool
            </p>
            <hr style="border-color:rgba(255,255,255,0.3); margin:1rem 0;">
            <p style="margin:0; font-size:0.95rem; opacity:0.75;">
                Civil Engineering Year 3 Continuous Assessment — Design of Steel Structures (2026)<br>
                <strong>Group 6</strong> &nbsp;|&nbsp; Makerere University Faculty of Engineering
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── About / Overview ─────────────────────────────────────────────────────
    st.markdown("## 📖 About This Tool")
    st.markdown(
        """
        The **Structural Engineering Toolkit** is an interactive, browser-based application built
        with [Streamlit](https://streamlit.io/) that brings together several steel-structure
        analysis and design routines in one place.  It was developed as a **Year 3 Continuous
        Assessment** for the *Design of Steel Structures* course, and is presented as Group 6's
        contribution to the wider civil engineering community.

        The tool covers everything from fundamental linear-algebra helpers (matrix multiplication,
        Gauss-Jordan elimination) up to full 2-D rigid **frame analysis** via the direct stiffness
        method, with Eurocode-3-aligned **steel section design** for tension members, compression
        struts, beams, and beam-columns.  Results can be downloaded as formatted Excel workbooks
        for use in design submissions or further calculation.
        """
    )

    st.divider()

    # ── Objectives ────────────────────────────────────────────────────────────
    col_obj, col_feat = st.columns(2, gap="large")

    with col_obj:
        st.markdown("### 🎯 Objectives")
        st.markdown(
            """
            - Provide a **unified, interactive** environment for common structural analysis tasks.
            - Implement and expose the **direct stiffness method** for 2-D truss and rigid-frame
              problems in an approachable UI.
            - Automate **Eurocode 3** steel section selection for tension, compression, beam, and
              beam-column design scenarios.
            - Generate **downloadable Excel design sheets** that can be used as evidence of
              compliance with design codes.
            - Serve as an educational resource for civil engineering students and practitioners.
            """
        )

    with col_feat:
        st.markdown("### ✨ Features")
        st.markdown(
            """
            | Module | What it does |
            |---|---|
            | **Matrix Multiplication** | Multiplies two arbitrarily-sized matrices; shows result and offers Excel export. |
            | **Gauss-Jordan Elimination** | Solves a system of linear equations via full row reduction with step-by-step output. |
            | **Truss Analysis & Design** | Reads joint/member Excel templates, assembles and solves the stiffness system, then selects tension and compression steel sections for every member. |
            | **Single Truss Member Design** | Designs one tension or compression member given a force and length. |
            | **Simple Beam Design** | Selects a UB section for a restrained or unrestrained beam given M, V, and span. |
            | **Beam Analysis & Design** | Analyses a multi-support beam (point loads + UDLs) via `indeterminatebeam`, extracts peak M & V, then designs the section. |
            | **Beam-Column Design** | Checks and designs a member subject to combined axial force and bending. |
            | **Frame Analysis & Design** | Assembles and solves a 2-D rigid frame by the direct stiffness method; shows member forces, node displacements, and designs all members to Eurocode 3. |
            """
        )

    st.divider()

    # ── How to Use ────────────────────────────────────────────────────────────
    st.markdown("## 🛠 How to Use")

    with st.expander("**1 — Matrix Multiplication & Gauss-Jordan**", expanded=False):
        st.markdown(
            """
            1. Select the task from the sidebar.
            2. Type your matrix (or augmented matrix) into the text box using the format shown
               — rows separated by `;`, columns by `,` (e.g. `1,2;3,4`).
            3. Click **Multiply** / **Solve**.
            4. Download the result with the **📥 Download Results Sheet** button.
            """
        )

    with st.expander("**2 — Truss Analysis & Design**", expanded=False):
        st.markdown(
            """
            1. Download the **joints** and **members** Excel templates provided on the task page.
            2. Fill in joint coordinates, support conditions, and applied loads in `joints.xlsx`.
            3. Fill in connectivity, member properties (cross-section, elastic modulus) in `members.xlsx`.
            4. Upload both files, configure the **Steel Properties** panel (grade, section shape,
               jointing type), and click **Run Analysis & Design**.
            5. Review the Tension / Compression design summary tables and cross-section diagrams,
               then download the full results workbook.
            """
        )

    with st.expander("**3 — Single Truss Member Design**", expanded=False):
        st.markdown(
            """
            1. Choose this task from the sidebar.
            2. Set steel grade and section preferences in the Steel Properties panel.
            3. Enter the axial force (kN), select Tension or Compression, and (for compression)
               the member length (mm).
            4. Click **Design Member** to see the chosen section, utilisation, and section diagram.
            """
        )

    with st.expander("**4 — Simple Beam Design**", expanded=False):
        st.markdown(
            """
            1. Choose steel grade.
            2. Enter the design bending moment (kNm), shear force (kN), and beam span (mm).
            3. Select **Restrained** or **Unrestrained**; for unrestrained beams also choose the
               beam condition and restraint level.
            4. Click **Design Beam** — the app picks the lightest satisfactory UB section and
               shows its cross-section diagram.
            """
        )

    with st.expander("**5 — Beam Analysis & Design**", expanded=False):
        st.markdown(
            """
            1. Set the steel grade, beam condition, and whether lateral restraint is provided.
            2. Enter the beam length and define supports (position and type: Pinned / Roller / Fixed).
            3. Add loads (Point Loads and/or UDLs).
            4. Click **Analyze & Design Beam** — the tool solves the beam, reports max M and V,
               selects a UB section, and produces an Excel report.
            """
        )

    with st.expander("**6 — Beam-Column Design**", expanded=False):
        st.markdown(
            """
            1. Enter the axial compression force (kN) and the design bending moment (kNm).
            2. Specify the member length and effective-length conditions.
            3. Click **Design Beam-Column** to obtain the recommended UC section and interaction
               check results.
            """
        )

    with st.expander("**7 — Frame Analysis & Design**", expanded=False):
        st.markdown(
            """
            1. Use the interactive tables on the task page to enter:
               - **Nodes** — ID, X-coordinate (m), Y-coordinate (m).
               - **Members** — ID, start node, end node, type (Beam / Column).
               - **Supports** — node ID and support type (Fixed / Pinned / Roller H / Roller V).
               - **Node Loads** — Fx (kN), Fy (kN), Mz (kNm) at any node.
               - **UDL Loads** — wx and wy (kN/m) along any member.
            2. Choose the steel grade.
            3. Click **Run Frame Analysis & Design** — the app solves the structure, displays member
               forces and node displacements, then designs each member.
            """
        )

    st.divider()

    # ── Libraries Used ────────────────────────────────────────────────────────
    st.markdown("## 📦 Libraries & Technologies Used")

    lib_col1, lib_col2, lib_col3 = st.columns(3, gap="medium")

    with lib_col1:
        st.markdown(
            """
            **🌐 Streamlit**
            *v≥1.0*

            Powers the entire web front-end — widgets, layouts, file uploaders, download buttons,
            and real-time interactivity without any JavaScript.

            ---

            **🔢 NumPy**
            *Scientific computing*

            Matrix assembly, inversion, eigenvalue decomposition, and all numerical linear-algebra
            operations underpinning the stiffness method.
            """
        )

    with lib_col2:
        st.markdown(
            """
            **📊 Pandas**
            *Data analysis*

            Structures and displays design-summary tables; used internally when reading Excel
            section databases and writing output workbooks.

            ---

            **📁 openpyxl**
            *Excel I/O*

            Reads the joints/members input templates and writes formatted Excel design reports
            for download.
            """
        )

    with lib_col3:
        st.markdown(
            """
            **📐 Matplotlib**
            *Visualisation*

            Renders annotated cross-section diagrams (I-sections, CHS, angle sections) that
            appear alongside every design result.

            ---

            **🏗 indeterminatebeam**
            *Beam analysis*

            Solves statically indeterminate beams — extracts support reactions, bending-moment
            diagrams, and shear envelopes for the *Beam Analysis & Design* module.
            """
        )

    st.divider()

    # ── Limitations ───────────────────────────────────────────────────────────
    st.markdown("## ⚠️ Limitations")
    st.warning(
        """
        Please read these limitations before using the tool for any structural design work:

        - **2-D only:** All structural analysis (truss and frame) is confined to a single vertical
          plane. Out-of-plane effects, torsion, and biaxial bending are not considered.
        - **Linear elastic analysis:** The solver assumes small displacements and linear material
          behaviour. Second-order (P-Δ / P-δ) effects are ignored.
        - **Fixed section properties for frame analysis:** The frame-analysis engine uses
          representative mid-range I-section properties (UB 457×191×67 for beams,
          UC 254×254×73 for columns) during the stiffness assembly; actual selected sections may
          differ, so a design-iteration loop is not implemented.
        - **Eurocode 3 scope:** Design checks follow EC3 principles for common cases. Some
          specialist checks (e.g. web crippling, patch loading, Class 4 sections) are outside the
          current scope.
        - **No dynamic or seismic analysis:** The tool is limited to static loading.
        - **Section databases:** The available section libraries are UB, UC, CHS, and equal/unequal
          angles loaded from the bundled Excel files. If a required section is not in the database,
          no result will be returned.
        - **Units:** Inputs and outputs must follow the conventions stated in each module
          (kN, kNm, m, mm as labelled). Mixing units will produce incorrect results.
        - **Gauss-Jordan solver:** The custom solver in `libfunc.py` requires non-zero diagonal
          pivots and does not implement partial pivoting; it will fail for some well-posed systems
          that need row swapping.
        """
    )

    st.divider()

    # ── Team ──────────────────────────────────────────────────────────────────
    st.markdown("## 👥 The Team — Group 6")
    st.markdown(
        """
        This tool was developed as a **Year 3 Continuous Assessment** for the
        *Design of Steel Structures* course (2026) by the following students:
        """
    )

    team = [
        {
            "name": "Nsamba Twalha Imran",
            "role": "System Designer & Lead Developer",
            "note": "Designed and built the full application architecture, analysis engines, and Streamlit interface.",
        },
        {
            "name": "Samuel Kalibala",
            "role": "Group Member",
            "note": "Contributed to the continuous assessment and verification of design outputs.",
        },
        {
            "name": "Natude Daniel",
            "role": "Group Member",
            "note": "Contributed to the continuous assessment and verification of design outputs.",
        },
        {
            "name": "Mwesigwa Eria",
            "role": "Group Member",
            "note": "Contributed to the continuous assessment and verification of design outputs.",
        },
    ]

    team_cols = st.columns(4, gap="medium")
    icons = ["👨‍💻", "👷", "👷", "👷"]

    for col, member, icon in zip(team_cols, team, icons):
        with col:
            st.markdown(
                f"""
                <div style="
                    border: 1px solid #dde3ea;
                    border-radius: 10px;
                    padding: 1.1rem 0.9rem;
                    text-align: center;
                    background: #f7faff;
                    height: 100%;
                ">
                    <div style="font-size:2.2rem;">{icon}</div>
                    <h4 style="margin:0.4rem 0 0.2rem 0; font-size:0.95rem; color:#1a3a5c;">
                        {member['name']}
                    </h4>
                    <p style="margin:0; font-size:0.78rem; color:#2e6da4; font-weight:600;">
                        {member['role']}
                    </p>
                    <p style="margin:0.5rem 0 0 0; font-size:0.75rem; color:#555;">
                        {member['note']}
                    </p>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.divider()

    # ── Footer ────────────────────────────────────────────────────────────────
    st.markdown(
        """
        <div style="text-align:center; color:#888; font-size:0.85rem; padding: 0.5rem 0 1rem 0;">
            🏛️ &nbsp; Presented as Group 6's contribution to the Civil Engineering community &nbsp; | &nbsp;
            Design of Steel Structures · 2026 &nbsp; | &nbsp;
            Built with ❤️ using Python &amp; Streamlit
        </div>
        """,
        unsafe_allow_html=True,
    )

# =====================================================
# MATRIX MULTIPLICATION
# =====================================================

elif task == "Matrix Multiplication":

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

                report_bytes = report_generator.matrix_mult_report(arg1, arg2, result)
                st.download_button(
                    label="📥 Download Results Sheet",
                    data=report_bytes,
                    file_name="matrix_multiplication_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
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
            aug_copy = arg.copy()
            solution = libfunc.solver(arg)

            st.success("Solution:")
            st.write(solution)

            report_bytes = report_generator.gauss_jordan_report(aug_copy)
            st.download_button(
                label="📥 Download Results Sheet",
                data=report_bytes,
                file_name="gauss_jordan_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

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

            else:
                # Build & offer results sheet — re-use global state already set by run_analysis_and_design_table
                _NSC, _NDOF = truss_analysis.assign_structure_coordinates()
                _S = truss_analysis.generate_stiffness_matrix(_NSC, _NDOF)
                _P = truss_analysis.form_load_vector(_NSC, _NDOF)
                _D = truss_analysis.solve_displacements(_S, _P)
                _mforces = truss_analysis.calculate_member_forces(_NSC, _D)

                report_bytes = report_generator.truss_report(
                    member_forces=_mforces,
                    tension_table=tension_table,
                    compression_table=compression_table,
                    grade=truss_analysis.grade,
                    shapeten=truss_analysis.shapeten,
                    shapecomp=truss_analysis.shapecomp,
                    jointing=truss_analysis.jointing,
                    nh=truss_analysis.nh,
                    d=truss_analysis.d,
                    stag=truss_analysis.stag,
                    s=truss_analysis.s,
                    p=truss_analysis.p,
                    ngs=truss_analysis.ngs,
                )
                st.download_button(
                    label="📥 Download Results Sheet",
                    data=report_bytes,
                    file_name="truss_analysis_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # ---- Section Visualizations ----
                st.subheader("📐 Section Visualizations")
                vis_cols = st.columns(2)

                if tension_table is not None and not tension_table.empty:
                    # Pick a representative tension section from the first member
                    first_ten = tension_table.iloc[0]
                    rep_ten_section = truss_analysis.ten_designer(abs(float(first_ten["Force (kN)"])) * 1000)
                    if rep_ten_section is not None:
                        with vis_cols[0]:
                            st.markdown(f"**Tension ({truss_analysis.shapeten})**")
                            fig_t = section_visualizer.visualize_tension_section(
                                truss_analysis.shapeten, rep_ten_section,
                                grade=truss_analysis.grade,
                                info_text=f"Shape: {first_ten['Size']}"
                            )
                            if fig_t is not None:
                                st.pyplot(fig_t, use_container_width=False)
                                plt.close(fig_t)

                if compression_table is not None and not compression_table.empty:
                    first_comp = compression_table.iloc[0]
                    _mlens = truss_analysis.calculate_member_lengths()
                    # find member id from label "Member X"
                    try:
                        _mid = int(str(first_comp["Member"]).split()[-1])
                        _clen = _mlens.get(_mid, 3000.0)
                    except Exception:
                        _clen = 3000.0
                    rep_comp_section = truss_analysis.comp_designer(
                        abs(float(first_comp["Force (kN)"])) * 1000, _clen
                    )
                    if rep_comp_section is not None:
                        col_idx = 1 if (tension_table is not None and not tension_table.empty) else 0
                        with vis_cols[col_idx]:
                            st.markdown(f"**Compression ({truss_analysis.shapecomp})**")
                            fig_c = section_visualizer.visualize_compression_section(
                                truss_analysis.shapecomp, rep_comp_section,
                                grade=truss_analysis.grade,
                                info_text=f"Shape: {first_comp['Size']}"
                            )
                            if fig_c is not None:
                                st.pyplot(fig_c, use_container_width=False)
                                plt.close(fig_c)

        except Exception as e:

            st.error(f"Analysis error: {e}")

elif task == "Single Truss Member Design":

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

                    report_bytes = report_generator.tension_design_report(
                        force_N=force * 1000,
                        grade=truss_analysis.grade,
                        shapeten=truss_analysis.shapeten,
                        section=section,
                        jointing=truss_analysis.jointing,
                        nh=truss_analysis.nh,
                        d=truss_analysis.d,
                        stag=truss_analysis.stag,
                        s=truss_analysis.s,
                        p=truss_analysis.p,
                        ngs=truss_analysis.ngs,
                    )
                    st.download_button(
                        label="📥 Download Results Sheet",
                        data=report_bytes,
                        file_name="tension_design_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    _show_section_figure(
                        section_visualizer.visualize_tension_section(
                            truss_analysis.shapeten, section,
                            grade=truss_analysis.grade,
                            info_text=f"Axial force: {force:.1f} kN (Tension)"
                        )
                    )

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

                    report_bytes = report_generator.compression_design_report(
                        force_N=force * 1000,
                        L=length,
                        grade=truss_analysis.grade,
                        shapecomp=truss_analysis.shapecomp,
                        section=section,
                    )
                    st.download_button(
                        label="📥 Download Results Sheet",
                        data=report_bytes,
                        file_name="compression_design_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    _show_section_figure(
                        section_visualizer.visualize_compression_section(
                            truss_analysis.shapecomp, section,
                            grade=truss_analysis.grade,
                            info_text=(
                                f"Axial force: {force:.1f} kN (Compression) | "
                                f"χ = {section['chi']:.3f} | "
                                f"Utilization: {round((force / (section['NbRd'] / 1000)) * 100, 1):.1f}%"
                            )
                        )
                    )

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

            if beam_type == "Restrained":
                report_bytes = report_generator.restrained_beam_report(
                    M=M, V=V,
                    grade=truss_analysis.grade,
                    result=result,
                )
            else:
                report_bytes = report_generator.unrestrained_beam_report(
                    M=M, V=V, L=L,
                    grade=truss_analysis.grade,
                    condition=truss_analysis.condition,
                    endcondition=truss_analysis.endcondition,
                    result=result,
                )
            st.download_button(
                label="📥 Download Results Sheet",
                data=report_bytes,
                file_name="beam_design_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            _show_section_figure(
                section_visualizer.visualize_beam_section(
                    result["Size"],
                    grade=truss_analysis.grade,
                    beam_type=beam_type + " Beam",
                    info_text=(
                        f"M = {M:.1f} kNm | V = {V:.1f} kN | "
                        f"Utilization: {result.get('Utilization (%)', '—')}%"
                    )
                )
            )

        except Exception as e:
            st.error(f"Error: {e}")

elif task == "Beam Analysis & Design":

    st.header("Beam Analysis & Design")

    steel_input_ui(True)
    condition = st.selectbox(
        "Beam Condition",
        ["Rolled", "Welded"]
    )
    latrestrain = st.checkbox("Provide Lateral Restraint")

    st.write("---")

    # =========================
    # BEAM LENGTH
    # =========================
    L = st.number_input("Beam Length (m)", value=10.0)

    # =========================
    # SUPPORTS
    # =========================
    st.subheader("Supports")

    n_supports = st.number_input("Number of Supports", min_value=2, value=2)

    supports = []

    for i in range(int(n_supports)):
        col1, col2 = st.columns(2)

        with col1:
            x = st.number_input(f"Support {i+1} Position (m)", key=f"sx_new{i}")

        with col2:
            typ = st.selectbox(
                f"Support {i+1} Type",
                ["Pinned", "Roller", "Fixed"],
                key=f"st_new{i}"
            )

        supports.append((x, typ.lower()))

    st.write("---")

    # =========================
    # LOADS
    # =========================
    st.subheader("Loads")

    n_loads = st.number_input("Number of Loads", min_value=1, value=1)

    loads = []

    for i in range(int(n_loads)):

        ltype = st.selectbox(
            f"Load {i+1} Type",
            ["Point Load", "UDL"],
            key=f"lt_new{i}"
        )

        if ltype == "Point Load":
            P = st.number_input(f"P{i+1} (kN)", key=f"P_new{i}")
            x = st.number_input(f"Position (m)", key=f"Px_new{i}")

            loads.append(("point", P, x))

        else:
            w = st.number_input(f"w{i+1} (kN/m)", key=f"w_new{i}")
            a = st.number_input(f"Start (m)", key=f"a_new{i}")
            b = st.number_input(f"End (m)", key=f"b_new{i}")

            loads.append(("udl", w, a, b))

    st.write("---")

    # =========================
    # RUN ANALYSIS
    # =========================
    if st.button("Analyze & Design Beam"):

        try:
            # -------------------------
            # CREATE BEAM
            # -------------------------
            beam = Beam(L)

            # -------------------------
            # ADD SUPPORTS
            # -------------------------
            for x, typ in supports:

                if typ == "pinned":
                    beam.add_supports(Support(x, (1, 1, 0)))

                elif typ == "roller":
                    beam.add_supports(Support(x, (0, 1, 0)))

                elif typ == "fixed":
                    beam.add_supports(Support(x, (1, 1, 1)))

                else:
                    raise ValueError(f"Unknown support type: {typ}")

            # -------------------------
            # ADD LOADS (IMPORTANT: convert to N)
            # -------------------------
            for load in loads:

                if load[0] == "point":
                    _, P, x = load
                    beam.add_loads(PointLoad(-P * 1000, x, 90))  # kN → N

                elif load[0] == "udl":
                    _, w, a, b = load
                    beam.add_loads(UDL(-1*w * 1000, (a, b), 90))  # kN/m → N/m

            # -------------------------
            # SOLVE
            # -------------------------
            beam.analyse()

            # -------------------------
            # EXTRACT RESULTS (CORRECT WAY)
            # -------------------------
            M = beam.get_bending_moment(return_absmax=True) / 1000  # → kNm
            V = beam.get_shear_force(return_absmax=True) / 1000     # → kN

            # plt.figure()
            # fig = beam.plot_beam_diagram()
            # st.pyplot(plt.gcf())
            # plt.close()

            st.info(f"Max Moment = {round(M,2)} kNm")
            st.info(f"Max Shear = {round(V,2)} kN")

            # # -------------------------
            # # OPTIONAL DEBUG INFO
            # # -------------------------
            # st.text("Beam Summary:")
            # st.text(beam)

            # reactions = {
            #     s[0]: beam.get_reaction(s[0])
            #     for s in supports
            # }
            # st.write("Support Reactions:", reactions)

            # -------------------------
            # AUTO RESTRAINT DETECTION
            # -------------------------
            types = [s[1] for s in supports]

            if all(t == "fixed" for t in types):
                restraint = "Full"

            elif "fixed" in types:
                restraint = "Partial"

            elif len(types) == 1 and types[0] == "fixed":
                restraint = "Cantilever"

            else:
                restraint = "Free"

            # -------------------------
            # DESIGN
            # -------------------------
            if latrestrain:
                result = truss_analysis.restrained_beam(M, V)

            else:
                truss_analysis.condition = condition
                truss_analysis.endcondition = restraint
                result = truss_analysis.unrestrained_beam(M, V, L * 1000)

            st.success("Design Result")
            st.dataframe(pd.DataFrame([result]))

            if latrestrain:
                report_bytes = report_generator.restrained_beam_report(
                    M=M, V=V,
                    grade=truss_analysis.grade,
                    result=result,
                )
            else:
                report_bytes = report_generator.unrestrained_beam_report(
                    M=M, V=V, L=L * 1000,
                    grade=truss_analysis.grade,
                    condition=condition,
                    endcondition=restraint,
                    result=result,
                )
            st.download_button(
                label="📥 Download Results Sheet",
                data=report_bytes,
                file_name="beam_analysis_design_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            beam_label = "Restrained Beam" if latrestrain else "Unrestrained Beam"
            _show_section_figure(
                section_visualizer.visualize_beam_section(
                    result["Size"],
                    grade=truss_analysis.grade,
                    beam_type=beam_label,
                    info_text=(
                        f"M = {round(M, 2):.2f} kNm | V = {round(V, 2):.2f} kN | "
                        f"Span = {L:.1f} m"
                    )
                )
            )

        except Exception as e:
            st.error(f"Error: {e}")

# =====================================================
# BEAM-COLUMN DESIGN
# =====================================================

elif task == "Beam-Column Design":

    st.header("Beam-Column Design (Axial + Bending)")
    steel_input_ui(True)

    shape = st.selectbox(
        "Shape",
        ["UB","UC"]
    )
    truss_analysis.condition = st.selectbox(
        "Steel Condition",
        ["Rolled","Welded"]
    )

    st.write("---")
    # -------------------------
    # User Inputs
    # -------------------------

    Ned = st.number_input(
        "Axial Force (kN)",
        value=100.0
    )*1000

    st.write("---")

    st.markdown("### About z-axis")

    col1, col2 = st.columns(2)

    with col1:
        M1_z = st.number_input("M1,z (kNm)", key="M1z", value = 100.0)*1e6

    with col2:
        M2_z = st.number_input("M2,z (kNm)", key="M2z", value = 0.02)*1e6

    # --- 3D toggle
    is_3D = st.checkbox("Enable 3D bending (y-axis moments)")

    # --- Y-axis moments (only if 3D is enabled)
    if is_3D:
        Myed = st.number_input(
            "Design Moment about Y (kNm)",
            value=100.0
        )*1e6
        # st.markdown("### About y-axis")

        # col3, col4 = st.columns(2)
        #
        # with col3:
        #     M1_y = st.number_input("M1,y (kNm)", key="M1y")*1e6
        #
        # with col4:
        #     M2_y = st.number_input("M2,y (kNm)", key="M2y")*1e6
        # Myed = max(M1_y, M2_y)
        # posseidony = Myed/min(M1_y, M2_y)
    else:
        Myed = 0.0
        # posseidony = 0.0
    Mzed = max(M1_z, M2_z)


    L = st.number_input(
        "Member Length (mm)",
        value=3000.0
    )
    # lat = st.checkbox("Laterally Restrained?")
    all_axis_diff = st.checkbox("Are restraints different for both axes?")
    if all_axis_diff:
        all_axis_similar = False
        truss_analysis.endcondition = []
        col1, col2 = st.columns(2)
        with col1:
            truss_analysis.endcondition.append(st.selectbox(
                "Fixity about Z-axis",
                ["Fixed-Fixed", "Fixed-Pinned","Pinned-Pinned","Fixed-Free"]
            ))
        with col2:
            truss_analysis.endcondition.append(st.selectbox(
                "Fixity about Y-axis",
                ["Fixed-Fixed", "Fixed-Pinned", "Pinned-Pinned", "Fixed-Free"]
            ))
    else:
        all_axis_similar = True
        truss_analysis.endcondition = st.selectbox(
            "Fixity about Z-axis",
            ["Fixed-Fixed", "Fixed-Pinned", "Pinned-Pinned", "Fixed-Free"]
        )
    st.write("---")

    # -------------------------
    # RUN DESIGN
    # -------------------------
    if st.button("Design Beam-Column 🚀"):

        # -------------------------
        # SAFETY CHECKS
        # -------------------------
        try:
            if M1_z == 0 or M2_z == 0:
                st.error("Moments M1,z and M2,z must both be non-zero.")
                st.stop()

            if min(M1_z, M2_z) == 0:
                st.error("Cannot compute moment ratio (division by zero).")
                st.stop()

            # Recompute safely
            Mzed = max(M1_z, M2_z)
            posseidon = Mzed / min(M1_z, M2_z)

            # Clamp to avoid extreme nonsense
            posseidon = np.clip(posseidon, 1.0, 10.0)

            C1 = 1 / (0.3 + 0.7 * (posseidon**2))

            # -------------------------
            # CALL DESIGN FUNCTION
            # -------------------------
            result = truss_analysis.beam_column(
                L=L,
                Ned=Ned,
                Mzed=Mzed,
                Myed=Myed,
                shape=shape,
                C1=C1,
                all_axis_similar=all_axis_similar
            )

            # -------------------------
            # DISPLAY RESULTS
            # -------------------------
            st.success("✅ Suitable section found!")

            st.write("### 📊 Results")

            st.write(f"**Section Designation:** {result['Designation']}")
            st.write(f"**Cross-section Class:** {result['class']}")

            st.write("---")

            st.write("### 🧱 Resistances")

            st.write(f"N_b,Rd = {result['N_b_Rd']/1000:.2f} kN")
            st.write(f"M_b,Rd = {result['M_b_Rd']/1e6:.2f} kNm")
            st.write(f"M_z,Rd = {result['M_z_Rd']/1e6:.2f} kNm")

            st.write("---")

            st.write("### ⚙️ Stability Factors")

            st.write(f"χ_y = {result['chi_y']:.3f}")
            st.write(f"χ_z = {result['chi_z']:.3f}")
            st.write(f"χ_LT = {result['chi_LT']:.3f}")

            st.write("---")

            st.write("### 🎯 Utilisation")

            U = result["utilisation"]

            if U <= 1:
                st.success(f"✅ Utilisation = {U:.3f} (SAFE)")
            else:
                st.error(f"❌ Utilisation = {U:.3f} (FAIL)")

            report_bytes = report_generator.beam_column_report(
                L=L,
                Ned=Ned,
                Mzed=Mzed,
                Myed=Myed,
                shape=shape,
                C1=C1,
                grade=truss_analysis.grade,
                endcondition=truss_analysis.endcondition,
                all_axis_similar=all_axis_similar,
                result=result,
            )
            st.download_button(
                label="📥 Download Results Sheet",
                data=report_bytes,
                file_name="beam_column_design_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            _show_section_figure(
                section_visualizer.visualize_beam_column_section(
                    result["Designation"],
                    shape=shape,
                    grade=truss_analysis.grade,
                    info_text=(
                        f"N = {Ned/1000:.1f} kN | M_z = {Mzed/1e6:.1f} kNm | "
                        f"Utilisation = {U:.3f} | Class {result['class']}"
                    )
                )
            )

        except Exception as e:
            st.error(f"❌ Error: {str(e)}")

# =====================================================
# FRAME ANALYSIS & DESIGN
# =====================================================

elif task == "Frame Analysis & Design":

    st.header("Frame Analysis & Design")

    st.info(
        "Enter your 2-D steel frame below using the tables. "
        "The tool performs a linear-elastic analysis (direct stiffness method, 3 DOF/node) "
        "and then designs each beam (UB) and column (UC) to Eurocode 3."
    )

    # ── Steel properties ──────────────────────────────────────────────────
    steel_input_ui(True)
    truss_analysis.condition = st.selectbox("Column Steel Condition", ["Rolled", "Welded"])
    beam_condition = st.selectbox(
        "Beam Lateral Condition",
        ["Restrained", "Unrestrained"],
        help="Restrained = full lateral restraint (M_pl,Rd). "
             "Unrestrained = LTB check required (M_b,Rd)."
    )
    if beam_condition == "Unrestrained":
        unrestrained_end = st.selectbox(
            "Beam End Condition (for LTB effective length)",
            ["Free", "Partial", "Full", "Cantilever"]
        )
    else:
        unrestrained_end = "Free"

    col_endcondition = st.selectbox(
        "Column End Condition (for buckling effective length)",
        ["Pinned-Pinned", "Fixed-Pinned", "Fixed-Fixed", "Fixed-Free"],
        help="Applied to all columns. Pinned-Pinned is conservative."
    )

    st.write("---")

    # ── Default data for tables ───────────────────────────────────────────
    _default_nodes = pd.DataFrame({
        "Node": [1, 2, 3, 4],
        "X (m)": [0.0, 6.0, 6.0, 0.0],
        "Y (m)": [0.0, 0.0, 4.0, 4.0],
    })

    _default_members = pd.DataFrame({
        "Member": [1, 2, 3],
        "Start Node": [1, 4, 2],
        "End Node":   [4, 3, 3],
        "Type":       ["Column", "Beam", "Column"],
    })

    _default_supports = pd.DataFrame({
        "Node":      [1, 2],
        "Condition": ["Fixed", "Fixed"],
    })

    _default_node_loads = pd.DataFrame({
        "Node":     [4],
        "Fx (kN)":  [15.0],
        "Fy (kN)":  [0.0],
        "Mz (kNm)": [0.0],
    })

    _default_udl = pd.DataFrame({
        "Member":    [2],
        "wx (kN/m)": [0.0],
        "wy (kN/m)": [-25.0],
    })

    # ── Nodes ─────────────────────────────────────────────────────────────
    st.subheader("① Nodes")
    st.caption("One row per node. Coordinates in metres (origin at bottom-left is conventional).")
    nodes_df = st.data_editor(
        _default_nodes,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Node":   st.column_config.NumberColumn("Node ID", min_value=1, step=1, format="%d"),
            "X (m)":  st.column_config.NumberColumn("X  (m)", format="%.3f"),
            "Y (m)":  st.column_config.NumberColumn("Y  (m)", format="%.3f"),
        },
        key="frame_nodes"
    )

    # ── Members ───────────────────────────────────────────────────────────
    st.subheader("② Members")
    st.caption(
        "Connect nodes to form the frame. "
        "Choose **Beam** (horizontal / inclined, designed as UB) or "
        "**Column** (vertical, designed as UC with axial + bending). "
        "**Beam** members that carry significant axial force alongside bending "
        "are automatically reclassified as **Beam-Columns** and designed to "
        "EC3 §6.3.3 (N+M interaction check) using UB sections."
    )
    members_df = st.data_editor(
        _default_members,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Member":     st.column_config.NumberColumn("Member ID", min_value=1, step=1, format="%d"),
            "Start Node": st.column_config.NumberColumn("Start Node", min_value=1, step=1, format="%d"),
            "End Node":   st.column_config.NumberColumn("End Node",   min_value=1, step=1, format="%d"),
            "Type":       st.column_config.SelectboxColumn("Type", options=["Beam", "Column"]),
        },
        key="frame_members"
    )

    # ── Supports ──────────────────────────────────────────────────────────
    st.subheader("③ Support Conditions")
    supports_df = st.data_editor(
        _default_supports,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Node":      st.column_config.NumberColumn("Node ID", min_value=1, step=1, format="%d"),
            "Condition": st.column_config.SelectboxColumn(
                "Condition",
                options=["Fixed", "Pinned", "Roller (H)", "Roller (V)"]
            ),
        },
        key="frame_supports"
    )

    st.write("---")

    # ── Loads ──────────────────────────────────────────────────────────────
    st.subheader("④ Loads")

    col_l1, col_l2 = st.columns(2)

    with col_l1:
        st.markdown("**Nodal Loads**")
        st.caption("Point forces / moments applied directly at nodes.")
        node_loads_df = st.data_editor(
            _default_node_loads,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Node":     st.column_config.NumberColumn("Node ID", min_value=1, step=1, format="%d"),
                "Fx (kN)":  st.column_config.NumberColumn("Fx  (kN)", format="%.2f"),
                "Fy (kN)":  st.column_config.NumberColumn("Fy  (kN)", format="%.2f"),
                "Mz (kNm)": st.column_config.NumberColumn("Mz  (kNm)", format="%.2f"),
            },
            key="frame_node_loads"
        )

    with col_l2:
        st.markdown("**Member Distributed Loads (UDL)**")
        st.caption(
            "Uniform loads along member length. "
            "wy is vertical (+ve = upward). wy = −25 kN/m means 25 kN/m downward."
        )
        udl_df = st.data_editor(
            _default_udl,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Member":    st.column_config.NumberColumn("Member ID", min_value=1, step=1, format="%d"),
                "wx (kN/m)": st.column_config.NumberColumn("wx  (kN/m)", format="%.2f"),
                "wy (kN/m)": st.column_config.NumberColumn("wy  (kN/m)", format="%.2f"),
            },
            key="frame_udl"
        )

    st.write("---")

    # ── Run Analysis & Design ──────────────────────────────────────────────
    if st.button("Analyze & Design Frame 🚀", type="primary"):

        try:
            # ── Convert DataFrames to lists ─────────────────────────────
            nodes_list   = nodes_df.dropna(subset=["Node"]).values.tolist()
            members_list = members_df.dropna(subset=["Member"]).values.tolist()
            supports_list = supports_df.dropna(subset=["Node"]).values.tolist()
            node_loads_list = node_loads_df.dropna(subset=["Node"]).values.tolist()
            udl_list        = udl_df.dropna(subset=["Member"]).values.tolist()

            if len(nodes_list) < 2:
                st.error("Please define at least 2 nodes.")
                st.stop()
            if len(members_list) < 1:
                st.error("Please define at least 1 member.")
                st.stop()
            if len(supports_list) < 1:
                st.error("Please define at least 1 support.")
                st.stop()

            # ── Run FEA ─────────────────────────────────────────────────
            member_results, node_disp = fa.analyse_frame(
                nodes=nodes_list,
                members=members_list,
                supports=supports_list,
                node_loads=node_loads_list,
                udl_loads=udl_list,
            )

            # ── Set globals required by truss_analysis functions ─────────
            truss_analysis.endcondition = col_endcondition

            member_design = {}
            design_errors = {}
            member_effective_types = {}

            # Thresholds for beam-column detection: a Beam member with axial
            # force and bending moment above these limits is redesigned as a
            # beam-column so that the N+M interaction is properly checked.
            _BC_N_THRESHOLD = 1.0   # kN
            _BC_M_THRESHOLD = 1.0   # kNm

            for mid, res in member_results.items():
                L_m   = res["length"]
                L_mm  = L_m * 1000
                N_kN  = max(abs(res["N_start"]), abs(res["N_end"]))
                V_kN  = max(abs(res["V_start"]), abs(res["V_end"]))
                M_kNm = max(abs(res["M_start"]), abs(res["M_end"]))

                # A Beam member that carries significant axial force alongside
                # bending must be treated as a beam-column.
                is_beam_column = (
                    res["type"] == "Beam"
                    and N_kN > _BC_N_THRESHOLD
                    and M_kNm > _BC_M_THRESHOLD
                )

                try:
                    if is_beam_column:
                        # Design as beam-column using UB shape (EC3 §6.3.3)
                        member_effective_types[mid] = "Beam-Column"
                        truss_analysis.endcondition = col_endcondition
                        N_N   = N_kN  * 1000
                        M_Nmm = M_kNm * 1e6
                        M1 = abs(res["M_start"])
                        M2 = abs(res["M_end"])
                        Mmax = max(M1, M2, 1e-6)
                        Mmin = max(min(M1, M2), 1e-6)
                        posseidon = np.clip(Mmax / Mmin, 1.0, 10.0)
                        C1 = 1.0 / (0.3 + 0.7 * posseidon ** 2)
                        dr = truss_analysis.beam_column(
                            L=L_mm, Ned=N_N, Mzed=M_Nmm, Myed=0.0,
                            shape="UB", C1=C1, all_axis_similar=True,
                        )
                        member_design[mid] = dr

                    elif res["type"] == "Beam":
                        member_effective_types[mid] = "Beam"
                        if beam_condition == "Restrained":
                            dr = truss_analysis.restrained_beam(M_kNm, V_kN)
                        else:
                            truss_analysis.endcondition = unrestrained_end
                            dr = truss_analysis.unrestrained_beam(M_kNm, V_kN, L_mm)
                        member_design[mid] = dr

                    else:  # Column
                        member_effective_types[mid] = "Column"
                        truss_analysis.endcondition = col_endcondition
                        N_N   = N_kN  * 1000
                        M_Nmm = M_kNm * 1e6
                        M1 = abs(res["M_start"])
                        M2 = abs(res["M_end"])
                        Mmax = max(M1, M2, 1e-6)
                        Mmin = max(min(M1, M2), 1e-6)
                        posseidon = np.clip(Mmax / Mmin, 1.0, 10.0)
                        C1 = 1.0 / (0.3 + 0.7 * posseidon ** 2)
                        dr = truss_analysis.beam_column(
                            L=L_mm, Ned=N_N, Mzed=M_Nmm, Myed=0.0,
                            shape="UC", C1=C1, all_axis_similar=True,
                        )
                        member_design[mid] = dr

                except Exception as exc:
                    design_errors[mid] = str(exc)
                    member_effective_types.setdefault(mid, res["type"])

            # ── Store everything in session state so it survives re-runs
            st.session_state["frame_results"] = {
                "member_results": member_results,
                "member_design":  member_design,
                "design_errors":  design_errors,
                "member_effective_types": member_effective_types,
                "nodes_list":     nodes_list,
                "members_list":   members_list,
                "supports_list":  supports_list,
                "node_loads_list": node_loads_list,
                "udl_list":       udl_list,
                "grade":          truss_analysis.grade,
                "beam_condition": beam_condition,
                "col_condition":  truss_analysis.condition,
                "col_endcondition": col_endcondition,
            }
            # Reset section-view toggles for a fresh run
            st.session_state["frame_view_sections"] = {}

        except Exception as e:
            st.error(f"❌ Error: {str(e)}")

    # ── Display persisted results (survives "View Section" button clicks) ──
    if "frame_results" in st.session_state:
        fr = st.session_state["frame_results"]
        member_results = fr["member_results"]
        member_design  = fr["member_design"]
        design_errors  = fr["design_errors"]
        member_effective_types = fr.get("member_effective_types", {})
        nodes_list     = fr["nodes_list"]
        members_list   = fr["members_list"]
        supports_list  = fr["supports_list"]
        node_loads_list = fr["node_loads_list"]
        udl_list       = fr["udl_list"]
        _grade         = fr["grade"]
        _beam_cond     = fr["beam_condition"]
        _col_cond      = fr["col_condition"]
        _col_ec        = fr["col_endcondition"]

        if "frame_view_sections" not in st.session_state:
            st.session_state["frame_view_sections"] = {}

        st.success("✅ Analysis complete!")

        st.subheader("📊 Member End Forces (from elastic analysis)")
        ana_rows = []
        for mid, res in member_results.items():
            N_max = max(abs(res["N_start"]), abs(res["N_end"]))
            V_max = max(abs(res["V_start"]), abs(res["V_end"]))
            M_max = max(abs(res["M_start"]), abs(res["M_end"]))
            ana_rows.append({
                "Member": mid,
                "Type":   res["type"],
                "Length (m)": round(res["length"], 3),
                "N_start (kN)": round(res["N_start"], 2),
                "N_end (kN)":   round(res["N_end"], 2),
                "V_start (kN)": round(res["V_start"], 2),
                "V_end (kN)":   round(res["V_end"], 2),
                "M_start (kNm)": round(res["M_start"], 2),
                "M_end (kNm)":   round(res["M_end"], 2),
                "Max |N| (kN)":  round(N_max, 2),
                "Max |V| (kN)":  round(V_max, 2),
                "Max |M| (kNm)": round(M_max, 2),
            })
        st.dataframe(pd.DataFrame(ana_rows), use_container_width=True)

        st.write("---")
        st.subheader("🔩 Design Results")

        for mid, res in member_results.items():
            mtype = res["type"]
            eff_type = member_effective_types.get(mid, mtype)

            with st.container():
                header_col, btn_col = st.columns([5, 1])

                if mid in design_errors:
                    with header_col:
                        st.error(
                            f"**Member {mid}** ({mtype}) — Design failed: "
                            f"{design_errors[mid]}"
                        )
                    continue

                dr = member_design[mid]

                if eff_type == "Beam":
                    size  = dr.get("Size", "—")
                    util  = dr.get("Utilization (%)", 0.0)
                    ok    = util <= 100.0
                    Mrd   = dr.get("M_Rd (kNm)", 0.0)
                    Vrd   = dr.get("V_Rd (kN)", 0.0)
                    M_Ed  = max(abs(res["M_start"]), abs(res["M_end"]))
                    V_Ed  = max(abs(res["V_start"]), abs(res["V_end"]))
                    row_d = {
                        "Member": mid, "Type": "Beam (UB)", "Section": size,
                        "M_Ed (kNm)": round(M_Ed, 2), "M_Rd (kNm)": round(Mrd, 2),
                        "V_Ed (kN)":  round(V_Ed, 2), "V_Rd (kN)":  round(Vrd, 2),
                        "Utilisation (%)": round(util, 1),
                        "Status": "PASS ✓" if ok else "FAIL ✗",
                    }
                    util_str = f"{util:.1f}%"

                elif eff_type == "Beam-Column":
                    size  = dr.get("Designation", "—")
                    U     = dr.get("utilisation", 0.0)
                    ok    = U <= 1.0
                    Nbrd  = dr.get("N_b_Rd", 0.0) / 1000
                    N_Ed  = max(abs(res["N_start"]), abs(res["N_end"]))
                    M_Ed  = max(abs(res["M_start"]), abs(res["M_end"]))
                    row_d = {
                        "Member": mid,
                        "Type": "Beam-Column (UB) ⚠️ reclassified",
                        "Section": size,
                        "N_Ed (kN)": round(N_Ed, 2), "N_b,Rd (kN)": round(Nbrd, 2),
                        "M_Ed (kNm)": round(M_Ed, 2),
                        "Class": dr.get("class", "—"),
                        "χ_y":  round(dr.get("chi_y",  0), 3),
                        "χ_z":  round(dr.get("chi_z",  0), 3),
                        "χ_LT": round(dr.get("chi_LT", 0), 3),
                        "Utilisation": round(U, 3),
                        "Status": "PASS ✓" if ok else "FAIL ✗",
                    }
                    util_str = f"{U:.3f}"

                else:  # Column
                    size  = dr.get("Designation", "—")
                    U     = dr.get("utilisation", 0.0)
                    ok    = U <= 1.0
                    Nbrd  = dr.get("N_b_Rd", 0.0) / 1000
                    N_Ed  = max(abs(res["N_start"]), abs(res["N_end"]))
                    M_Ed  = max(abs(res["M_start"]), abs(res["M_end"]))
                    row_d = {
                        "Member": mid, "Type": "Column (UC)", "Section": size,
                        "N_Ed (kN)": round(N_Ed, 2), "N_b,Rd (kN)": round(Nbrd, 2),
                        "M_Ed (kNm)": round(M_Ed, 2),
                        "Class": dr.get("class", "—"),
                        "χ_y":  round(dr.get("chi_y",  0), 3),
                        "χ_z":  round(dr.get("chi_z",  0), 3),
                        "χ_LT": round(dr.get("chi_LT", 0), 3),
                        "Utilisation": round(U, 3),
                        "Status": "PASS ✓" if ok else "FAIL ✗",
                    }
                    util_str = f"{U:.3f}"

                with header_col:
                    label = eff_type if eff_type != "Beam-Column" else f"{mtype} → Beam-Column"
                    if ok:
                        st.success(
                            f"**Member {mid}** ({label}) — **{size}** — "
                            f"Utilisation: {util_str}"
                        )
                    else:
                        st.error(
                            f"**Member {mid}** ({label}) — **{size}** — "
                            f"FAIL — Utilisation: {util_str}"
                        )
                    if eff_type == "Beam-Column":
                        st.info(
                            "ℹ️ This member carries both significant axial force and "
                            "bending moment and has been automatically reclassified as a "
                            "**Beam-Column** and designed to EC3 §6.3.3 (N+M interaction)."
                        )
                    st.dataframe(pd.DataFrame([row_d]), use_container_width=True)

                with btn_col:
                    st.write("")
                    st.write("")
                    if st.button("📐 View\nSection", key=f"view_frame_{mid}"):
                        current = st.session_state["frame_view_sections"].get(mid, False)
                        st.session_state["frame_view_sections"][mid] = not current

                # Show section figure if toggled on
                if st.session_state["frame_view_sections"].get(mid, False):
                    info = (
                        f"M_Ed = {round(max(abs(res['M_start']), abs(res['M_end'])), 2):.2f} kNm  |  "
                        f"V_Ed = {round(max(abs(res['V_start']), abs(res['V_end'])), 2):.2f} kN  |  "
                        f"N_Ed = {round(max(abs(res['N_start']), abs(res['N_end'])), 2):.2f} kN"
                    )
                    if eff_type == "Beam":
                        fig = section_visualizer.visualize_beam_section(
                            size, grade=_grade,
                            beam_type=f"{_beam_cond} Beam",
                            info_text=info,
                        )
                    elif eff_type == "Beam-Column":
                        fig = section_visualizer.visualize_beam_column_section(
                            size, shape="UB", grade=_grade, info_text=info,
                        )
                    else:
                        fig = section_visualizer.visualize_beam_column_section(
                            size, shape="UC", grade=_grade, info_text=info,
                        )
                    _show_section_figure(fig, header=f"📐 Member {mid} — {size}")

        # ── Download results sheet ────────────────────────────────────────
        st.write("---")
        if member_design:
            report_bytes = report_generator.frame_design_report(
                grade=_grade,
                beam_condition=_beam_cond,
                col_condition=_col_cond,
                col_endcondition=_col_ec,
                member_analysis=member_results,
                member_design=member_design,
                member_effective_types=member_effective_types,
                nodes=nodes_list,
                members=members_list,
                supports=supports_list,
                node_loads=node_loads_list,
                udl_loads=udl_list,
            )
            st.download_button(
                label="📥 Download Results Sheet",
                data=report_bytes,
                file_name="frame_analysis_design_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )



