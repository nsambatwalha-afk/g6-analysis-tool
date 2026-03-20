import streamlit as st
import numpy as np
import pandas as pd
import libfunc
import truss_analysis

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

    st.header("Beam Analysis & Design")

    steel_input_ui(True)

    st.write("---")

    # -------------------------
    # GEOMETRY INPUT
    # -------------------------

    beam_length = st.number_input("Total Beam Length (m)", value=10.0)

    st.subheader("Supports")

    n_supports = st.number_input("Number of Supports", min_value=2, value=2)

    supports = []

    for i in range(int(n_supports)):
        col1, col2 = st.columns(2)

        with col1:
            x = st.number_input(f"Support {i+1} Position (m)", key=f"sx{i}")
        with col2:
            typ = st.selectbox(
                f"Support {i+1} Type",
                ["Pinned", "Fixed", "Roller"],
                key=f"st{i}"
            )

        supports.append((x, typ))

    # -------------------------
    # LOAD INPUT
    # -------------------------

    st.subheader("Loads")

    n_loads = st.number_input("Number of Loads", min_value=1, value=1)

    loads = []

    for i in range(int(n_loads)):

        load_type = st.selectbox(
            f"Load {i+1} Type",
            ["Point Load", "UDL"],
            key=f"lt{i}"
        )

        if load_type == "Point Load":
            P = st.number_input(f"P{i+1} (kN)", key=f"P{i}")
            x = st.number_input(f"Position (m)", key=f"Px{i}")

            loads.append(("point", P, x))

        else:
            w = st.number_input(f"w{i+1} (kN/m)", key=f"w{i}")
            a = st.number_input(f"Start (m)", key=f"a{i}")
            b = st.number_input(f"End (m)", key=f"b{i}")

            loads.append(("udl", w, a, b))

    st.write("---")

    # -------------------------
    # BEAM ANALYSIS (STIFFNESS)
    # -------------------------

    def beam_analysis(L_total, supports, loads):
        import numpy as np

        # 1. GENERATE NODES: Include supports AND point load positions
        point_load_pos = [l[2] for l in loads if l[0] == "point"]
        node_positions = sorted(set([0, L_total] + [s[0] for s in supports] + point_load_pos))

        n = len(node_positions)
        dof = 2 * n
        K = np.zeros((dof, dof))
        F_nodal = np.zeros(dof)  # For direct point loads
        F_fef = np.zeros(dof)  # For Fixed End Forces (UDLs)

        E = 210000000  # kN/m2
        I = 0.0001  # m4 (Standard UB approx)
        EI = E * I

        # 2. ASSEMBLY
        for i in range(n - 1):
            x1, x2 = node_positions[i], node_positions[i + 1]
            L = x2 - x1
            idx = [2 * i, 2 * i + 1, 2 * (i + 1), 2 * (i + 1) + 1]

            # Element Stiffness Matrix (Units: kN and m)
            k_el = (EI / L ** 3) * np.array([
                [12, 6 * L, -12, 6 * L],
                [6 * L, 4 * L ** 2, -6 * L, 2 * L ** 2],
                [-12, -6 * L, 12, -6 * L],
                [6 * L, 2 * L ** 2, -6 * L, 4 * L ** 2]
            ])
            for a in range(4):
                for b in range(4):
                    K[idx[a], idx[b]] += k_el[a, b]

            # 3. UDL LOAD MAPPING (Fixed End Forces)
            for load in loads:
                if load[0] == "udl":
                    _, w, a, b = load
                    # Calculate overlap of UDL on this specific element
                    overlap_start = max(x1, a)
                    overlap_end = min(x2, b)

                    if overlap_end > overlap_start:
                        curr_L = overlap_end - overlap_start
                        # Simplified: Assume UDL covers full element if it touches it
                        # For CEDAT precision, use exact FEF integration if UDL is partial
                        f_fef = np.array([w * L / 2, w * L ** 2 / 12, w * L / 2, -w * L ** 2 / 12])
                        F_fef[idx] += f_fef

        # 4. POINT LOAD MAPPING
        for load in loads:
            if load[0] == "point":
                _, P, x = load
                node_idx = node_positions.index(x)
                F_nodal[2 * node_idx] += P  # Note: Convention down is positive or negative?
                # Consistent with FEF: Downward load is positive in F vector

        F_total = F_nodal + F_fef

        # 5. BOUNDARY CONDITIONS
        fixed_dofs = []
        for x, typ in supports:
            node_idx = node_positions.index(x)
            fixed_dofs.append(2 * node_idx)  # Vertical translation fixed
            if typ == "Fixed":
                fixed_dofs.append(2 * node_idx + 1)  # Rotation fixed

        free_dofs = [i for i in range(dof) if i not in fixed_dofs]
        d = np.zeros(dof)
        if len(free_dofs) > 0:
            d[free_dofs] = np.linalg.solve(K[np.ix_(free_dofs, free_dofs)], F_total[free_dofs])

        # 6. RECOVER MAX M & V (Corrected)
        Mmax = 0
        Vmax = 0

        for i in range(n - 1):
            x1, x2 = node_positions[i], node_positions[i + 1]
            L = (x2 - x1) * 1000  # mm
            idx = [2 * i, 2 * i + 1, 2 * (i + 1), 2 * (i + 1) + 1]

            # 1. Forces from nodal displacements (d)
            k_local = (EI / L ** 3) * np.array([
                [12, 6 * L, -12, 6 * L],
                [6 * L, 4 * L ** 2, -6 * L, 2 * L ** 2],
                [-12, -6 * L, 12, -6 * L],
                [6 * L, 2 * L ** 2, -6 * L, 4 * L ** 2]
            ])
            f_disp = k_local @ d[idx]

            # 2. Subtract the Fixed End Forces for this specific element
            f_fef = np.zeros(4)
            for load in loads:
                if load[0] == "udl":
                    _, w, a, b = load
                    # Check if UDL covers this segment
                    if x1 >= a and x2 <= b:
                        wN = w  # N/mm
                        f_fef = np.array([wN * L / 2, wN * L ** 2 / 12, wN * L / 2, -wN * L ** 2 / 12])

            # 3. Total Internal Forces
            # The internal forces are the displacement forces MINUS the fixed end forces
            f_internal = f_disp - f_fef

            # Convert back to kN and kNm for comparison
            V_start = abs(f_internal[0]) / 1000
            M_start = abs(f_internal[1]) / 1e6
            V_end = abs(f_internal[2]) / 1000
            M_end = abs(f_internal[3]) / 1e6

            Vmax = max(Vmax, V_start, V_end)
            Mmax = max(Mmax, M_start, M_end)

        return Mmax, Vmax
    # -------------------------
    # AUTO RESTRAINT
    # -------------------------

    def detect_restraint(supports):
        types = [s[1] for s in supports]

        if all(t == "Fixed" for t in types):
            return "Full"
        elif "Fixed" in types:
            return "Partial"
        elif types[0] == "Fixed" and len(types) == 1:
            return "Cantilever"
        else:
            return "Free"

    # -------------------------
    # RUN
    # -------------------------

    if st.button("Analyze & Design Beam"):

        try:
            M, V = beam_analysis(beam_length, supports, loads)

            st.info(f"Max Moment = {round(M,2)} kNm")
            st.info(f"Max Shear = {round(V,2)} kN")

            restraint = detect_restraint(supports)

            if restraint == "Full":
                result = truss_analysis.restrained_beam(M, V)
            else:
                truss_analysis.condition = "Rolled"
                truss_analysis.endcondition = restraint
                result = truss_analysis.unrestrained_beam(M, V, beam_length*1000)

            st.success("Design Result")
            st.dataframe(pd.DataFrame([result]))

        except Exception as e:
            st.error(f"Error: {e}")