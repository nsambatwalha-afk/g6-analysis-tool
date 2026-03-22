
import openpyxl
import numpy as np
import pandas as pd

PJ = 0
JP = 0
COORD = 0
MSUP = 0
MPRP = 0
EM = 0
CP = 0
NCJT = 2  # number of coordinates per joint (2 for trusses)
grade = "SXXX"
shapecomp = "XXX"
shapeten = "XXX"
jointing = "XXXX"
stag = False
nh = 0.0
d = 0.0
s = 0.0
p = 0.0
ngs = 0.0
condition = ""
endcondition = ""


def inputxl(joints_file, members_file):
    global PJ, JP, COORD, MSUP, MPRP, EM, CP

    if joints_file is None or members_file is None:
        raise ValueError("Missing joints or members Excel file.")

    try:
        joints_wb = openpyxl.load_workbook(joints_file)
        members_wb = openpyxl.load_workbook(members_file)
    except Exception as e:
        raise ValueError(f"Error loading Excel files: {e}")

    joints = joints_wb.active
    members = members_wb.active

    memberno = int(members.cell(row=4, column=21).value)
    jointno = int(joints.cell(row=2, column=21).value)

    coord = [
        [float(joints.cell(row=i, column=2).value),
         float(joints.cell(row=i, column=3).value)]
        for i in range(2, jointno + 2)
    ]
    COORD = np.array(coord)

    msup = []
    for i in range(2, jointno + 2):
        if float(joints.cell(row=i, column=4).value) > 0.0:
            msup.append([
                float(joints.cell(row=i, column=1).value),
                float(joints.cell(row=i, column=9).value),
                float(joints.cell(row=i, column=10).value)
            ])
    MSUP = np.array(msup)

    mprp = [
        [float(members.cell(row=i, column=c).value) for c in range(2, 6)]
        for i in range(2, memberno + 2)
    ]
    MPRP = np.array(mprp)

    NMP = int(members.cell(row=2, column=21).value)
    EM = np.array([[float(members.cell(row=i, column=8).value)]
                   for i in range(2, NMP + 2)])

    NCP = int(members.cell(row=3, column=21).value)
    CP = np.array([[float(members.cell(row=i, column=9).value)]
                   for i in range(2, NCP + 2)])

    jp, pj = [], []
    for i in range(2, jointno + 2):
        if float(joints.cell(row=i, column=6).value) > 0.0:
            jp.append([float(joints.cell(row=i, column=1).value)])
            pj.append([
                float(joints.cell(row=i, column=7).value),
                float(joints.cell(row=i, column=8).value)
            ])

    JP = np.array(jp)
    PJ = np.array(pj)

def set_steel_properties(
    grade_in,
    shapeten_in,
    shapecomp_in,
    jointing_in,
    nh_in=0,
    d_in=0,
    staggered=False,
    s_in=0,
    p_in=0,
    ngs_in=0
):
    global grade, shapeten, shapecomp, jointing
    global stag, nh, d, s, p, ngs

    grade = grade_in
    shapeten = shapeten_in
    shapecomp = shapecomp_in
    jointing = jointing_in

    if jointing == "bolt":
        stag = staggered
        nh = nh_in
        d = d_in

        if stag:
            s = s_in
            p = p_in
            ngs = ngs_in
        else:
            s, p, ngs = 0, 0, 0

    elif jointing == "weld":
        stag = False
        nh, d, s, p, ngs = 0, 0, 0, 0, 0

    else:
        raise ValueError("Jointing must be 'bolt' or 'weld'")


def table_reader(table, val):
    if table == "CHS":
        ex = openpyxl.load_workbook("CHS.xlsx")
        ex = ex.active
        n = 0.0
        i = 1
        while n<val:
            i = i + 1
            n = float(ex.cell(row=i, column=4).value)
        dia = float(ex.cell(row=i, column=1).value)
        thickness = float(ex.cell(row=i, column=2).value)
        area = n
        rg = float(ex.cell(row=i, column=7).value)*10.0
        return area, dia, thickness, rg
    elif table == "Angle":
        ex = openpyxl.load_workbook("Angle.xlsx")
        ex = ex.active
        n = 0.0
        i = 44
        while n<=val:
            i = i - 1
            n = float(ex.cell(row=i, column=6).value)
        size = ex.cell(row=i, column=1).value
        thickness = float(ex.cell(row=i, column=2).value)
        area = n
        return area, size, thickness
    elif table == "UC":
        ex = openpyxl.load_workbook("UC-2.xlsx").active
        n = 0.0
        i = 1
        while n<val:
            i = i + 1
            n = float(ex.cell(row=i, column=14).value)
        area = n
        size = ex.cell(row=i, column=1).value
        Iy = float(ex.cell(row=i, column=2).value)*10000.0
        Iz = float(ex.cell(row=i, column=3).value)*10000.0
        ry = float(ex.cell(row=i, column=4).value)*10.0
        rz = float(ex.cell(row=i, column=5).value)*10.0
        if Iy<Iz:
            I = Iy
            R = ry
            axis = "y"
        else:
            axis = "z"
            I = Iz
            R = rz
        return area, size, I, R, axis
    elif table == "UB":
        ex = openpyxl.load_workbook("UB-2.xlsx").active
        n = 0.0
        i = 1
        while n < val:
            i = i + 1
            n = float(ex.cell(row=i, column=14).value)
        area = n
        size = ex.cell(row=i, column=1).value
        Iy = float(ex.cell(row=i, column=2).value) * 10000.0
        Iz = float(ex.cell(row=i, column=3).value) * 10000.0
        ry = float(ex.cell(row=i, column=4).value) * 10.0
        rz = float(ex.cell(row=i, column=5).value) * 10.0
        if Iy < Iz:
            I = Iy
            R = ry
            axis = "y"
        else:
            axis = "z"
            I = Iz
            R = rz
        return area, size, I, R, axis

    else:
        print("\n\nTable not recognized\n\n")
        return None

def ten_designer(f):
    global grade, shapeten, jointing, stag, nh, d, s, p, ngs
    grades = openpyxl.load_workbook("grades.xlsx").active
    i = 1
    pop = "sjbdcusdd"
    while pop!=grade and i<=7:
        i = i+1
        pop = grades.cell(row=i, column=1).value
    if i>7:
        print("\n\nGrade not recognized\n\n")
        return None
    fy = float(grades.cell(row=i, column=2).value)
    fu = float(grades.cell(row=i, column=3).value)
    areq = f/fy
    anet = 0.0
    agross = 0.0
    counter = 0
    while anet<areq and counter<=1000:
        counter = counter + 1
        section = table_reader(shapeten, max(agross,areq))
        if section is None:
            return None
        agross = section[0]
        t = section[2]
        if jointing == "weld":
            anet = agross
        elif jointing == "bolt" and not stag:
            anet = agross - nh * t * (d + 2.0)
        elif jointing == "bolt" and stag:
            anet = agross - nh * t * (d + 2.0) + ngs * (((s ** 2) * t) / 4 * p)
        else:
            print("\n\nJointing not recognized.\n\n")
    if counter==1000:
        print("\n\nInfinite loop in section sizing. - 1\n\n")
        return None
    npl = (agross*fy)/1.0
    nu = (0.9*anet*fu)/1.25
    counter = 1
    while f>min(npl,nu) and counter<=1000:
        counter = counter + 1
        section = table_reader(shapeten, agross)
        if section is None:
            return None
        agross = section[0]
        t = section[2]
        if jointing == "weld":
            anet = agross
        elif jointing == "bolt" and not stag:
            anet = agross - nh * t * (d + 2.0)
        elif jointing == "bolt" and stag:
            anet = agross - nh * t * (d + 2.0) + ngs * (((s ** 2) * t) / 4 * p)
        else:
            print("\n\nSomething went wrong. Error code 1\n\n")
            return None
        npl = (agross * fy) / 1.0
        nu = (0.9 * anet * fu) / 1.25
    if counter==1000:
        print("\n\nInfinite loop in section sizing. - 2\n\n")
        return None
    return section


def comp_designer(f,L):
    global grade, shapecomp, stag, nh, d, s, p, ngs
    grades = openpyxl.load_workbook("grades.xlsx").active
    i = 1
    pop = "sjbdcusdd"
    while pop != grade and i <= 7:
        i = i + 1
        pop = grades.cell(row=i, column=1).value
    if i > 7:
        print("Grade not recognized")
        return None
    fy = float(grades.cell(row=i, column=2).value)
    E = 210000  # MPa
    if shapecomp == "UC" or shapecomp == "UB":
        alpha = 0.34
    elif shapecomp == "CHS":
        alpha = 0.21
    elif shapecomp == "Angle":
        alpha = 0.49
    else:
        print("\n\nCompression shape not recognized\n\n")
        return None
    if shapecomp == "UB" or shapecomp == "UC":
        agross = f/fy
        counter = 0
        while counter <= 1000:
            counter += 1
            section = table_reader(shapecomp, agross)
            if section is None:
                return None
            A, size, I, r, axis = section
            Ncr = (np.pi ** 2 * E * I) / (L ** 2)
            lam_bar = np.sqrt((A * fy) / Ncr)
            phi = 0.5 * (1 + alpha * (lam_bar - 0.2) + lam_bar ** 2)
            chi = 1 / (phi + np.sqrt(phi ** 2 - lam_bar ** 2))
            NbRd = chi * A * fy
            if NbRd >= f:
                return {
                    "area": A,
                    "size": size,
                    "I": I,
                    "r": r,
                    "axis": axis,
                    "lambda_bar": lam_bar,
                    "chi": chi,
                    "NbRd": NbRd
                }
            else:
                agross = A
    elif shapecomp == "CHS":

        agross = f / fy
        counter = 0

        while counter <= 1000:
            counter += 1
            section = table_reader(shapecomp, agross)
            if section is None:
                return None
            A, dia, t, r = section

            # For CHS: I = A * r^2
            I = A * (r ** 2)
            Ncr = (np.pi ** 2 * E * I) / (L ** 2)
            lam_bar = np.sqrt((A * fy) / Ncr)
            phi = 0.5 * (1 + alpha * (lam_bar - 0.2) + lam_bar ** 2)
            chi = 1 / (phi + np.sqrt(max(phi ** 2 - lam_bar ** 2, 0)))
            NbRd = chi * A * fy
            if NbRd >= f:
                return {
                    "area": A,
                    "size": f"CHS {dia}x{t}",
                    "I": I,
                    "r": r,
                    "axis": "circular",
                    "lambda_bar": lam_bar,
                    "chi": chi,
                    "NbRd": NbRd
                }

            else:
                agross = A












def assign_structure_coordinates():
    global COORD, MSUP
    NJ = COORD.shape[0]
    NR = int(np.sum(MSUP[:, 1:]))
    NDOF = NCJT * NJ - NR
    NSC = np.zeros(NCJT * NJ, dtype=int)
    J = 0
    K = NDOF

    for i in range(NJ):
        is_support = np.where(MSUP[:, 0].astype(int) == i + 1)[0]
        if len(is_support) == 0:
            for j in range(NCJT):
                J += 1
                NSC[i * NCJT + j] = J
        else:
            sup_idx = is_support[0]
            for j in range(NCJT):
                if int(MSUP[sup_idx, j + 1]) == 1:
                    K += 1
                    NSC[i * NCJT + j] = K
                else:
                    J += 1
                    NSC[i * NCJT + j] = J
    return NSC, NDOF

def generate_stiffness_matrix(NSC, NDOF):
    global COORD, MPRP, EM, CP
    S = np.zeros((NDOF, NDOF))
    for m in range(MPRP.shape[0]):
        jb, je, mat_idx, cp_idx = map(lambda x: int(x) - 1, MPRP[m])
        E = EM[int(mat_idx), 0]
        A = CP[int(cp_idx), 0]
        x1, y1 = COORD[jb]
        x2, y2 = COORD[je]
        L = np.hypot(x2 - x1, y2 - y1)
        CX = (x2 - x1) / L
        CY = (y2 - y1) / L
        z = E * A / L
        c2, s2, cs = CX ** 2, CY ** 2, CX * CY
        k = z * np.array([[ c2, cs, -c2, -cs],
                          [ cs, s2, -cs, -s2],
                          [-c2, -cs, c2, cs],
                          [-cs, -s2, cs, s2]])
        idx = [NSC[jb * NCJT] - 1, NSC[jb * NCJT + 1] - 1,
               NSC[je * NCJT] - 1, NSC[je * NCJT + 1] - 1]
        for i in range(4):
            for j in range(4):
                if idx[i] < NDOF and idx[j] < NDOF:
                    S[idx[i], idx[j]] += k[i, j]
    return S

def calculate_member_lengths():
    global COORD, MPRP

    member_lengths = {}

    for m in range(MPRP.shape[0]):
        member_id = m + 1
        jb, je = int(MPRP[m,0]) - 1, int(MPRP[m,1]) - 1

        x1, y1 = COORD[jb]
        x2, y2 = COORD[je]

        L = np.hypot(x2 - x1, y2 - y1)

        member_lengths[member_id] = L

    return member_lengths

def form_load_vector(NSC, NDOF):
    global JP, PJ
    P = np.zeros(NDOF)
    for i in range(JP.shape[0]):
        joint = int(JP[i, 0]) - 1
        fx, fy = PJ[i]
        idx_x = NSC[joint * NCJT] - 1
        idx_y = NSC[joint * NCJT + 1] - 1
        if idx_x < NDOF: P[idx_x] = fx
        if idx_y < NDOF: P[idx_y] = fy
    return P

def solve_displacements(S, P):
    return np.linalg.solve(S, P)

def calculate_member_forces(NSC, D):
    global COORD, MPRP, EM, CP
    member_forces = {}

    for m in range(MPRP.shape[0]):
        member_id = m + 1
        jb, je, mat_idx, cp_idx = map(lambda x: int(x) - 1, MPRP[m])
        E = EM[int(mat_idx), 0]
        A = CP[int(cp_idx), 0]
        x1, y1 = COORD[jb]
        x2, y2 = COORD[je]
        L = np.hypot(x2 - x1, y2 - y1)
        CX = (x2 - x1) / L
        CY = (y2 - y1) / L

        idx = [NSC[jb * NCJT] - 1, NSC[jb * NCJT + 1] - 1,
               NSC[je * NCJT] - 1, NSC[je * NCJT + 1] - 1]

        u = np.array([D[i] if i < len(D) else 0.0 for i in idx])
        delta = CX * (u[2] - u[0]) + CY * (u[3] - u[1])
        axial_force = E * A / L * delta

        member_forces[member_id] = axial_force

    return member_forces


#def run_truss_analysis(verbose=False):
    inputxl()
    NSC, NDOF = assign_structure_coordinates()
    S = generate_stiffness_matrix(NSC, NDOF)
    P = form_load_vector(NSC, NDOF)
    D = solve_displacements(S, P)

    member_forces = calculate_member_forces(NSC, D)

    # Always show clean axial force table
    member_ids = [f"Member {m}" for m in member_forces.keys()]
    force_values = list(member_forces.values())

    force_table = pd.DataFrame({
        "Member": member_ids,
        "Axial Force": force_values
    })

    print("\nMember Forces:")
    print(force_table.to_string(index=False))

    # Only show detailed design output if requested
    if verbose:
        print("\nDetailed Design Results:")
        for m, res in design_results.items():
            print(f"Member {m}: {res}")

    return member_forces, design_results#

def design_single_member(force, member_type, length=1000):
    """
    Designs a single steel member.

    force: kN
    member_type: "Tension" or "Compression"
    length: mm (only needed for compression)
    """

    f = abs(force) * 1000  # convert to N

    if member_type == "Tension":

        section = ten_designer(f)

        if section is None:
            return None

        area = section[0]
        size = section[1]
        thickness = section[2]

        return {
            "Type": "Tension",
            "Shape": shapeten,
            "Size": size,
            "Thickness": thickness,
            "Area": area
        }

    elif member_type == "Compression":

        section = comp_designer(f, length)

        if section is None:
            return None

        return {
            "Type": "Compression",
            "Shape": shapecomp,
            "Size": section["size"],
            "Axis": section["axis"],
            "χ": round(section["chi"], 3),
            "NbRd (kN)": round(section["NbRd"]/1000, 3),
            "Capacity Utilization (%)": round((force/(section["NbRd"]/1000))*100, 3)
        }

    else:
        return None

def run_analysis_and_design_table(joints_file, members_file):

    if joints_file is None or members_file is None:
        raise ValueError("All three Excel files must be uploaded.")

    # ---- Structural Analysis ----
    inputxl(joints_file, members_file)

    NSC, NDOF = assign_structure_coordinates()

    S = generate_stiffness_matrix(NSC, NDOF)

    P = form_load_vector(NSC, NDOF)

    D = solve_displacements(S, P)

    member_forces = calculate_member_forces(NSC, D)

    # ---- Design containers ----
    tension_results = []
    compression_results = []

    member_lengths = calculate_member_lengths()

    for member_id, force in member_forces.items():

        L = member_lengths[member_id]

        if force > 0:

            section = ten_designer(force * 1000)

            if section is not None:

                area = section[0]
                size = section[1]
                thickness = section[2]

                tension_results.append({
                    "Member": f"Member {member_id}",
                    "Force (kN)": round(force, 3),
                    "Shape": shapeten,
                    "Size": size,
                    "Thickness": thickness
                })

        elif force < 0:

            section = comp_designer(-force * 1000, L)

            if section is not None:

                compression_results.append({
                    "Member": f"Member {member_id}",
                    "Force (kN)": round(force, 3),
                    "Shape": shapecomp,
                    "Size": section["size"],
                    "Axis": section["axis"],
                    "Capacity Utilization (%)": round((-force/(section["NbRd"]/1000))*100, 3),
                    "χ": round(section["chi"], 3),
                    "NbRd (kN)": round(section["NbRd"]/1000, 3)
                })

    tension_table = pd.DataFrame(tension_results)
    compression_table = pd.DataFrame(compression_results)

    return tension_table, compression_table

def restrained_beam(M, V):
    global grade

    grades = openpyxl.load_workbook("grades.xlsx").active

    # ---- Get fy ----
    for i in range(2, 10):
        if grades.cell(row=i, column=1).value == grade:
            fy = float(grades.cell(row=i, column=2).value)
            break

    gamma_M0 = 1.0

    ex = openpyxl.load_workbook("UB-2.xlsx").active

    i = 1

    while True:
        i += 1

        size = ex.cell(row=i, column=1).value

        if size is None:
            raise ValueError("No suitable section found in UB table.")

        Wyy = float(ex.cell(row=i, column=8).value) * 1000  # mm³
        Wzz = float(ex.cell(row=i, column=9).value) * 1000
        if Wyy > Wzz:
            axis = "y"
        else:
            axis = "z"
        Wpl = max(Wyy, Wzz)
        tw = float(ex.cell(row=i, column=15).value)           # web thickness
        h = float(ex.cell(row=i, column=16).value)            # depth
        tf = float(ex.cell(row=i, column=17).value)
        hw = h - 2 * tf
        Av = hw * tw  # mm²

        Aw = hw * tw

        # ---- SHEAR CHECK ----
        Vpl_Rd = (Av * fy) / (np.sqrt(3) * gamma_M0) / 1000  # kN

        if V > Vpl_Rd:
            continue  # FAIL

        # ---- MOMENT CAPACITY ----
        Mpl_Rd = (Wpl * fy) / gamma_M0 / 1e6  # kNm

        if V / Vpl_Rd <= 0.5:
            M_Rd = Mpl_Rd

        else:
            rho = (2 * V / Vpl_Rd - 1) ** 2

            M_Rd = (
                (Wpl - (rho * Aw**2) / (4 * tw)) * fy / gamma_M0
            ) / 1e6  # kNm

        # ---- FINAL CHECK ----
        if M <= M_Rd:
            return {
                "Type": "Restrained Beam",
                "Size": size,
                "M_Rd (kNm)": round(M_Rd, 2),
                "V_Rd (kN)": round(Vpl_Rd, 2),
                "Utilization (%)": round((M / M_Rd) * 100, 2)
            }

def unrestrained_beam(M, V, L):
    global condition, endcondition, grade

    # ---- Load steel grade ----
    grades = openpyxl.load_workbook("grades.xlsx").active

    for i in range(2, 10):
        if grades.cell(row=i, column=1).value == grade:
            fy = float(grades.cell(row=i, column=2).value)
            break

    if 'fy' not in locals():
        raise ValueError("Steel grade not found in grades.xlsx")

    E = 210000  # MPa
    gamma_M1 = 1.0

    # ---- Load UB table ----
    ex = openpyxl.load_workbook("UB-2.xlsx").active

    i = 1

    while True:
        i += 1

        size = ex.cell(row=i, column=1).value

        # ---- Stop if no more sections ----
        if size is None:
            raise ValueError("No suitable section found in UB table.")

        # ---- Section properties ----
        Wyy = float(ex.cell(row=i, column=8).value) * 1000  # mm³
        Wpl = Wyy  # Use major axis only

        Iy = float(ex.cell(row=i, column=2).value) * 10000  # mm⁴
        Iz = float(ex.cell(row=i, column=3).value) * 10000
        I = Iy

        tw = float(ex.cell(row=i, column=15).value)
        h = float(ex.cell(row=i, column=16).value)
        tf = float(ex.cell(row=i, column=17).value)
        b = float(ex.cell(row=i, column=19).value)
        iz = float(ex.cell(row=i, column=5).value)*10.0
        if endcondition=="Free":
            k = 1.0
        elif endcondition=="Partial":
            k = 0.85
        elif endcondition=="Full":
            k = 0.7
        elif endcondition=="Cantilever":
            k = 2.0
        else:
            raise ValueError("Endcondition not recognized.")
        lamz = (k*L)/iz
        laml = np.pi * np.sqrt(E/fy)
        lamzba = lamz / laml
        c1 = 1.0
        u = 0.9
        vee = 1.0
        bew = 1.0
        lamltb = (1/np.sqrt(c1))*u*vee*lamzba*np.sqrt(bew)
        hoverb = h/b
        if condition == "Rolled":
            if hoverb <= 2:
                alt = 0.34
            else:
                alt = 0.49
            phi = 0.5 * (1 + alt*(lamltb - 0.4) + 0.75*(lamltb**2))
            chi = min(1 / (phi + np.sqrt(phi**2 - 0.75*(lamltb**2))),1.0)
        else:
            if hoverb <= 2 and condition == "Welded":
                alt = 0.49
            elif hoverb > 2 and condition == "Welded":
                alt = 0.76
            else:
                alt = 0.76
            phi = 0.5 * (1 + alt*(lamltb - 0.2) + (lamltb**2))
            chi = min(1 / (phi + np.sqrt(phi**2 - (lamltb**2))),1.0)
        Mbrd = chi * Wpl * (fy/gamma_M1)
        # NOTE:
        # Mbrd is computed twice in this function.
        # The first computation (above) is in N·mm and is overwritten later by a kNm version.
        # This is redundant and potentially confusing.
        # TODO: Remove the earlier Mbrd definition and keep only the final (post-shear reduction) calculation.

        hw = h - 2 * tf
        Av = hw * tw  # shear area (mm²)
        Aw = hw * tw  # web area for reduction

        # ---- SHEAR CHECK ----
        Vpl_Rd = (Av * fy) / (np.sqrt(3) * gamma_M1) / 1000  # kN (Note: V is in kN)

        if V > Vpl_Rd:
            continue  # reject section

        # ---- HIGH SHEAR REDUCTION ----
        # If V is more than 50% of Vpl_Rd, reduce the moment capacity
        if V > 0.5 * Vpl_Rd:
            rho = ((2 * V / Vpl_Rd) - 1) ** 2
            # Reduced plastic modulus (Wpl_reduced)
            # Subtract the 'lost' capacity of the web area
            # Formula: M_y,V,Rd = (Wpl - (rho * Aw^2 / 4tw)) * fy / gamma_M0
            W_reduced = Wpl - (rho * (Aw ** 2)) / (4 * tw)
        else:
            W_reduced = Wpl

        # ---- UPDATED DESIGN MOMENT ----
        # Use W_reduced instead of Wpl in your Mbrd calculation
        # Note: Divide by 10^6 if W is mm3 and fy is MPa to get kNm
        Mbrd = (chi * W_reduced * fy / gamma_M1) / 1e6

        # ---- FINAL CHECK ----
        if M <= Mbrd:
            return {
                "Type": "Unrestrained Beam",
                "Size": size,
                "x_LT": round(chi, 3),
                "Mb_Rd (kNm)": round(Mbrd, 2),
                "Vpl_Rd": round(Vpl_Rd, 2),
                "Utilization (%)": round((M / Mbrd) * 100, 2)
            }

        # # ---- SHEAR REDUCTION ----
        # if V / Vpl_Rd <= 0.5:
        #      W_eff = Wpl
        # else:
        #      rho = (2 * V / Vpl_Rd - 1) ** 2
        #      W_eff = max(Wpl - (rho * Aw**2) / (4 * tw), 0.1 * Wpl)
        #
        # # ---- LTB (improved approximation) ----
        # Mcr = 2.5 * (np.pi**2 * E * I) / (L**2)
        #
        # lam = np.sqrt((W_eff * fy) / Mcr)
        #
        # alpha = 0.34
        # phi = 0.5 * (1 + alpha * (lam - 0.2) + lam**2)
        #
        # chi = 1 / (phi + np.sqrt(max(phi**2 - lam**2, 0)))
        #
        # # ---- DESIGN MOMENT ----
        # Mb_Rd = chi * W_eff * fy / gamma_M1 / 1e6  # kNm
        #
        # # ---- CHECK ----
        # if M <= Mb_Rd:
        #     return {
        #         "Type": "Unrestrained Beam",
        #         "Size": size,
        #         "χ_LT": round(chi, 3),
        #         "Mb_Rd (kNm)": round(Mb_Rd, 2),
        #         "Utilization (%)": round((M / Mb_Rd) * 100, 2)
        #     }



def section_class(fy, section, Ned):
    """
    Classifies an I/H section (UC/UB) to EC3 (Class 1–4)

    Parameters:
    H   : overall depth (mm)
    b   : flange width (mm)
    tf  : flange thickness (mm)
    tw  : web thickness (mm)
    fy  : yield strength (MPa)
    Ned : axial force (N) (compression +ve)
    A   : area (mm²)

    Returns:
    int → section class (1 to 4)
    """

    # -------------------------
    # BASIC PARAMETERS
    # -------------------------
    H = section.h
    b = section.b
    tf = section.tf
    tw = section.tw
    A = section.A
    eps = np.sqrt(235.0 / fy)

    d = H - 2.0 * tf                 # clear web depth
    c = (b - tw) / 2.0              # flange outstand

    # Avoid division issues
    if A * fy == 0:
        return 4

    alpha = 0.5 * (1.0 + Ned / (A * fy))

    # Clamp alpha to sensible EC3 range
    alpha = np.clip(alpha, 0.0, 1.0)

    # -------------------------
    # FLANGE CLASS
    # -------------------------
    lambda_f = c / tf

    if lambda_f <= 9.0 * eps:
        flange_class = 1
    elif lambda_f <= 10.0 * eps:
        flange_class = 2
    elif lambda_f <= 14.0 * eps:
        flange_class = 3
    else:
        flange_class = 4

    # -------------------------
    # WEB CLASS
    # -------------------------
    lambda_w = d / tw

    if alpha > 0.5:
        denom = (13.0 * alpha - 1.0)

        # Prevent divide-by-zero or negative weirdness
        if denom <= 0:
            return 4

        limit1 = 396.0 * eps / denom
        limit2 = 456.0 * eps / denom

    else:
        # Prevent division by zero
        if alpha <= 0:
            return 4

        limit1 = 36.0 * eps / alpha
        limit2 = 41.5 * eps / alpha

    if lambda_w <= limit1:
        web_class = 1
    elif lambda_w <= limit2:
        web_class = 2
    elif lambda_w <= 42.0 * eps:
        web_class = 3
    else:
        web_class = 4

    # -------------------------
    # FINAL CLASS
    # -------------------------
    return int(max(flange_class, web_class))

def eff_L(L, endcondition):
    if endcondition=="Pinned-Pinned":
        return L
    elif endcondition=="Fixed-Fixed":
        return L*0.5
    elif endcondition=="Fixed-Pinned":
        return L*0.7
    elif endcondition=="Fixed-Free":
        return L*2.0
    else:
        return None

class Section:
    def __init__(self, designation, A, Iy, Iz, Wpl, Wel, Iw, It, tw, h, tf, b, iz, seclass=None, alphay=None, alphaz=None):
        self.designation = designation
        self.A = A
        self.Iy = Iy
        self.Iz = Iz
        self.Wpl = Wpl
        self.Wel = Wel
        self.Iw = Iw
        self.It = It
        self.tw = tw
        self.h = h
        self.tf = tf
        self.b = b
        self.iz = iz
        self.seclass = seclass
        self.alphay = alphay
        self.alphaz = alphaz




def beam_column_table(shape, i):
    global condition, grade
    if shape == "UB":
        ex = openpyxl.load_workbook("UB-2.xlsx").active
        i = max(i, 17)
    elif shape == "UC":
        ex = openpyxl.load_workbook("UC-2.xlsx").active
    designation = ex.cell(row=i, column=1).value
    A = float(ex.cell(row=i, column=14).value)
    Iy = float(ex.cell(row=i, column=2).value) * 10000
    Iz = float(ex.cell(row=i, column=3).value) * 10000
    Wpl = float(ex.cell(row=i, column=8).value) * 1000
    Wel = float(ex.cell(row=i, column=6).value) * 1000
    Iw = float(ex.cell(row=i, column=12).value) * (100 ** 6)
    It = float(ex.cell(row=i, column=13).value) * 10000
    tw = float(ex.cell(row=i, column=15).value)
    h = float(ex.cell(row=i, column=16).value)
    tf = float(ex.cell(row=i, column=17).value)
    b = float(ex.cell(row=i, column=19).value)
    iz = float(ex.cell(row=i, column=5).value) * 10.0
    if condition == "Rolled":
        if (h/b)<=1.2 and not grade=="S460":
            if tf<=100:
                alphay = 0.34
                alphaz = 0.49
            else:
                alphay = 0.76
                alphaz = 0.76
        elif (h/b)>1.2 and not grade=="S460":
            if tf<=40:
                alphay = 0.21
                alphaz = 0.34
            else:
                alphay = 0.34
                alphaz = 0.49
        else:
            if tf<=40:
                alphay = 0.13
                alphaz = 0.13
            elif tf>100:
                alphay = 0.49
                alphaz = 0.49
            else:
                alphay = 0.21
                alphaz = 0.21
    else:
        if tf<=40:
            alphay = 0.34
            alphaz = 0.49
        else:
            alphay = 0.49
            alphaz = 0.76



    outer = Section(designation, A, Iy, Iz, Wpl, Wel, Iw, It, tw, h,tf, b, iz)
    outer.alphay = alphay
    outer.alphaz = alphaz
    return outer


def beam_column(L, Ned, Mzed, Myed, shape, C1, all_axis_similar=True):
    global condition, grade
    grades = openpyxl.load_workbook("grades.xlsx").active

    for i in range(2, 10):
        if grades.cell(row=i, column=1).value == grade:
            fy = float(grades.cell(row=i, column=2).value)
            break

    if 'fy' not in locals():
        raise ValueError("Steel grade not found in grades.xlsx")
    if all_axis_similar:
        Lcry = eff_L(L, endcondition)
        Lcrz = Lcry
    else:
        Lcrz = eff_L(L, endcondition[0])
        Lcry = eff_L(L, endcondition[1])
    E = 210000
    G = 81000
    i = 1
    if shape == "UB":
        ex = openpyxl.load_workbook("UB-2.xlsx").active
    elif shape == "UC":
        ex = openpyxl.load_workbook("UC-2.xlsx").active
    else:
        raise ValueError("Unknown shape")
    A0 = Ned/fy
    A = 0.0
    while True:
        while not A>A0:
            i += 1
            A = float(ex.cell(row=i, column=14).value)
        pop = beam_column_table(shape, i)
        pop.seclass = section_class(fy,pop,Ned)
        if pop.seclass == 1 or pop.seclass == 2:
            Nrd = (pop.A*fy)
            Mrd = pop.Wpl*fy
        elif pop.seclass == 3:
            Nrd = (pop.A*fy)
            Mrd = pop.Wel*fy
        elif pop.seclass == 4:
            raise ValueError("Unsupported class. Section is Class 4 which is outside our scope.")
        Ncry = ((np.pi**2)*E*pop.Iy)/(Lcry**2)
        Ncrz = ((np.pi**2)*E*pop.Iz)/(Lcrz**2)
        lamy = np.sqrt((pop.A*fy)/Ncry)
        lamz = np.sqrt((pop.A*fy)/Ncrz)
        phiy = 0.5 * (1+pop.alphay*(lamy-0.2)+(lamy**2))
        phiz = 0.5 * (1+pop.alphaz*(lamz-0.2)+(lamz**2))
        chiy = 1 / (phiy + np.sqrt(phiy**2 - lamy**2))
        chiz = 1 / (phiz + np.sqrt(phiz**2 - lamz**2))
        term2 = (pop.A*fy)
        chi = min(chiy, chiz)
        Nbrd = chi * term2
        # -------------------------
        # LATERAL TORSIONAL BUCKLING
        # -------------------------
        LcrLT = Lcry  # assumption (you can refine later)

        Mcr = C1 * ((np.pi**2 * E * pop.Iz) / (LcrLT**2)) * np.sqrt(
            (pop.Iw / pop.Iz) +
            ((LcrLT**2 * G * pop.It) / (np.pi**2 * E * pop.Iz))
        )

        # Select correct section modulus
        if pop.seclass in [1, 2]:
            Wy = pop.Wpl
        else:
            Wy = pop.Wel

        lamLT = np.sqrt((Wy * fy) / Mcr)

        # LTB imperfection factor (Eurocode typical for rolled I-sections)
        alpha_LT = 0.34

        phi_LT = 0.5 * (1 + alpha_LT * (lamLT - 0.2) + lamLT**2)
        chi_LT = 1 / (phi_LT + np.sqrt(phi_LT**2 - lamLT**2))

        Mbrd = chi_LT * Wy * fy

        # -------------------------
        # MINOR AXIS BENDING RESISTANCE
        # -------------------------
        if pop.seclass in [1, 2]:
            Wz = pop.Wpl  # approximation (strictly should be separate Wpl,z)
        else:
            Wz = pop.Wel

        Mzrd = Wz * fy

        # -------------------------
        # INTERACTION FACTORS (SIMPLIFIED EC3)
        # -------------------------
        # NOTE: This is a simplified safe approximation
        ny = Ned / Nbrd

        kyy = 1 + 0.6 * ny
        kzz = 1 + 0.6 * ny

        # Limit k to EC3 reasonable bounds
        kyy = max(1.0, kyy)
        kzz = max(1.0, kzz)

        # -------------------------
        # INTERACTION CHECK
        # -------------------------
        util_y = (Ned / Nbrd) + kyy * (Myed / Mbrd)
        util_z = (Ned / Nbrd) + kzz * (Mzed / Mzrd)

        U = max(util_y, util_z)

        # -------------------------
        # CHECK
        # -------------------------
        if U <= 1.0:
            return {
                "Designation": pop.designation,
                "class": pop.seclass,
                "N_b_Rd": Nbrd,
                "M_b_Rd": Mbrd,
                "M_z_Rd": Mzrd,
                "utilisation": U,
                "chi_y": chiy,
                "chi_z": chiz,
                "chi_LT": chi_LT
            }

        # Otherwise try next section
        i += 1

        # Safety break
        if i > ex.max_row:
            raise ValueError("No suitable section found.")




