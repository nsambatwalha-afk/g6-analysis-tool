
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

def inputxl_steel(steel_file):
    global grade, shapecomp, shapeten, jointing, stag, nh, d, s, p, ngs

    if steel_file is None:
        raise ValueError("Steel_entry.xlsx file missing.")

    try:
        data = openpyxl.load_workbook(steel_file).active
    except Exception as e:
        raise ValueError(f"Error reading Steel_entry.xlsx: {e}")

    grade = data.cell(row=3, column=2).value
    shapeten = data.cell(row=4, column=2).value
    shapecomp = data.cell(row=6, column=2).value
    jointing = data.cell(row=5, column=2).value

    if jointing == "bolt":
        stag1 = int(data.cell(row=10, column=2).value)

        if stag1 == 1:
            stag = True
            s = float(data.cell(row=15, column=2).value)
            p = float(data.cell(row=16, column=2).value)
            ngs = float(data.cell(row=17, column=2).value)

        elif stag1 == 0:
            stag = False

        else:
            raise ValueError("Invalid stagger input in Steel_entry.xlsx")

        nh = float(data.cell(row=11, column=2).value)
        d = float(data.cell(row=12, column=2).value)


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

    inputxl_steel()

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

def run_analysis_and_design_table(joints_file, members_file, steel_file):

    if joints_file is None or members_file is None or steel_file is None:
        raise ValueError("All three Excel files must be uploaded.")

    # ---- Structural Analysis ----
    inputxl(joints_file, members_file)

    NSC, NDOF = assign_structure_coordinates()

    S = generate_stiffness_matrix(NSC, NDOF)

    P = form_load_vector(NSC, NDOF)

    D = solve_displacements(S, P)

    member_forces = calculate_member_forces(NSC, D)

    # ---- Steel data ----
    inputxl_steel(steel_file)

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