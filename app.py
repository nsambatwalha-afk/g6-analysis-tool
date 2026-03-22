import streamlit as st
import numpy as np
import pandas as pd
import libfunc
import truss_analysis
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
        "Beam Analysis & Design",
        "Beam-Column Design"
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
    condition = st.selectbox(
        "Beam Condition",
        ["Rolled", "Welded"]
    )
    latrestrain = st.checkbox("Laterally Restrained?")

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
                    beam.add_loads(PointLoad(-P * 1000, x))  # kN → N

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
        M1_z = st.number_input("M1,z (kNm)", key="M1z")*1e6

    with col2:
        M2_z = st.number_input("M2,z (kNm)", key="M2z")*1e6

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

            st.write(f"**Section Index:** {result['section_index']}")
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

        except Exception as e:
            st.error(f"❌ Error: {str(e)}")


