
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


# ---------------------------------------------------------------------------
# Extended section databases – used when the required area exceeds the
# largest entry in the corresponding Excel file.
#
# CHS  – (area mm², diameter mm, thickness mm, radius of gyration mm)
#         sorted ascending by area; current xlsx max ≈ 24 700 mm²
# ---------------------------------------------------------------------------
_CHS_EXTENDED = [
    (29858, 610,  16, 210.1),   # CHS 610x16
    (33866, 559,  20, 190.7),   # CHS 559x20
    (37071, 610,  20, 208.7),   # CHS 610x20
    (40212, 660,  20, 226.4),   # CHS 660x20
    (43417, 711,  20, 244.4),   # CHS 711x20
    (46621, 762,  20, 262.4),   # CHS 762x20
    (49826, 813,  20, 280.5),   # CHS 813x20
    (56172, 914,  20, 316.2),   # CHS 914x20
    (69822, 914,  25, 314.4),   # CHS 914x25
    (77833, 1016, 25, 350.5),   # CHS 1016x25
]

# Angle – (area mm², size string, thickness mm)
#          sorted ascending by area; current xlsx max ≈ 9 060 mm²
_ANGLE_EXTENDED = [
    (9760,  "200x200", 26),  # EA 200x200x26
    (10460, "200x200", 28),  # EA 200x200x28
    (11900, "250x250", 25),  # EA 250x250x25
    (13260, "250x250", 28),  # EA 250x250x28
    (14990, "250x250", 32),  # EA 250x250x32
    (16330, "250x250", 35),  # EA 250x250x35
    (17100, "300x300", 30),  # EA 300x300x30
    (19800, "300x300", 35),  # EA 300x300x35
]

# UC – (area mm², designation, I_min mm⁴, r_min mm, axis)
#       sorted ascending by area; current xlsx max ≈ 80 800 mm²
#       I_min and r_min correspond to the weaker (z-z) axis for all UC sections.
_UC_EXTENDED = [
    (94800,  "356x406x744 UC",  1_180_000_000, 111.0, "z"),
    (114600, "356x406x900 UC",  1_470_000_000, 113.0, "z"),
    (138300, "356x406x1086 UC", 1_820_000_000, 115.0, "z"),
]

# UB – (area mm², designation, I_min mm⁴, r_min mm, axis)
#       sorted ascending by area; current xlsx max ≈ 20 600 mm²
#       I_min and r_min correspond to the weaker (z-z) axis for all UB sections.
_UB_EXTENDED = [
    (22000, "762x267x173 UB",  68_500_000,  55.8, "z"),   # 762x267x173
    (22800, "610x305x179 UB",  93_100_000,  63.9, "z"),   # 610x305x179
    (25100, "762x267x197 UB",  81_700_000,  57.0, "z"),   # 762x267x197
    (28000, "762x267x220 UB",  89_800_000,  56.6, "z"),   # 762x267x220
    (28500, "914x305x224 UB",  112_000_000, 62.7, "z"),   # 914x305x224
    (32000, "762x267x251 UB",  104_000_000, 57.2, "z"),   # 762x267x251
    (34600, "1016x305x272 UB", 145_000_000, 64.7, "z"),   # 1016x305x272
    (36800, "914x305x289 UB",  156_000_000, 65.2, "z"),   # 914x305x289
    (44500, "1016x305x349 UB", 185_000_000, 64.5, "z"),   # 1016x305x349
    (48500, "914x305x381 UB",  232_000_000, 69.2, "z"),   # 914x305x381
    (62000, "1016x305x487 UB", 317_000_000, 71.6, "z"),   # 1016x305x487
]


def _chs_extended(val):
    """Return the smallest extended CHS section with area >= val, or None."""
    for area, dia, t, rg in _CHS_EXTENDED:
        if area >= val:
            return area, dia, t, rg
    return None


def _angle_extended(val):
    """Return the smallest extended Angle section with area > val, or None.

    Uses strict inequality (>) to mirror the Angle loop's exit condition
    ``while n <= val`` (i.e. the xlsx loop stops at the first n that is
    *strictly greater* than val).  The other shape helpers use ``>=``
    because their loops use ``while n < val``.
    """
    for area, size, t in _ANGLE_EXTENDED:
        if area > val:
            return area, size, t
    return None


def _uc_extended(val):
    """Return the smallest extended UC section with area >= val, or None."""
    for area, size, I, R, axis in _UC_EXTENDED:
        if area >= val:
            return area, size, I, R, axis
    return None


def _ub_extended(val):
    """Return the smallest extended UB section with area >= val, or None."""
    for area, size, I, R, axis in _UB_EXTENDED:
        if area >= val:
            return area, size, I, R, axis
    return None


# ---------------------------------------------------------------------------
# Full extended section databases for beam_column, restrained_beam and
# unrestrained_beam.  These lists store all EC3 section properties needed by
# those functions so they can fall back gracefully when the xlsx tables are
# exhausted, matching the pattern introduced for table_reader in _UB_EXTENDED
# / _UC_EXTENDED.
#
# Tuple layout (mirrors UB-2 / UC-2 xlsx column order):
#   (designation, A mm², Iy cm⁴, Iz cm⁴, ry cm, rz cm,
#    Wel,y cm³, Wel,z cm³, Wpl,y cm³, Wpl,z cm³,
#    Iw dm⁶, It cm⁴, tw mm, h mm, tf mm, b mm)
#
# Properties estimated consistently with _UB_EXTENDED / _UC_EXTENDED:
#   Iz = A × rz²   (Iz values taken from _UB_EXTENDED / _UC_EXTENDED)
#   Iy = A × ry²   (ry per series: 762×267→27.7 cm, 610×305→24.6 cm,
#                   914×305→37.8 cm, 1016×305→40.0 cm)
#   Wel = 2Iy/h ;  Wpl ≈ 1.12 × Wel  (shape factor ≈ 1.12)
#   Iw  ≈ Iz × h² / 4  (cm⁴ × cm² = cm⁶, converted to dm⁶ for storage)
#   It  ≈ (2b tf³ + (h−2tf) tw³) / 3  (Saint-Venant torsion, mm⁴ → cm⁴)
# ---------------------------------------------------------------------------
_UB_FULL_EXTENDED = [
    # desig                 A      Iy      Iz    ry    rz  Wely  Welz   Wply  Wplz    Iw    It   tw    h   tf    b
    ("762x267x173 UB",  22000, 168804,  6850, 27.7, 5.58, 4431,  513,  4963,  770,  9.94,  255, 14,  762, 22, 267),
    ("610x305x179 UB",  22800, 137977,  9310, 24.6, 6.39, 4524,  610,  5067,  915,  8.66,  359, 13,  610, 25, 305),
    ("762x267x197 UB",  25100, 192590,  8170, 27.7, 5.70, 5055,  612,  5662,  918, 11.86,  430, 15,  762, 27, 267),
    ("762x267x220 UB",  28000, 214841,  8980, 27.7, 5.66, 5640,  673,  6317, 1010, 13.04,  626, 16,  762, 31, 267),
    ("914x305x224 UB",  28500, 407219, 11200, 37.8, 6.27, 8913,  734,  9983, 1101, 23.39,  399, 16,  914, 24, 305),
    ("762x267x251 UB",  32000, 245533, 10400, 27.7, 5.70, 6445,  779,  7219, 1169, 15.10, 1089, 17,  762, 38, 267),
    ("1016x305x272 UB", 34600, 553600, 14500, 40.0, 6.47,10898,  951, 12206, 1427, 37.42,  796, 16, 1016, 32, 305),
    ("914x305x289 UB",  36800, 525813, 15600, 37.8, 6.51,11507, 1023, 12888, 1535, 32.58, 1036, 18,  914, 35, 305),
    ("1016x305x349 UB", 44500, 712000, 18500, 40.0, 6.45,14016, 1213, 15698, 1820, 47.73, 1944, 19, 1016, 44, 305),
    ("914x305x381 UB",  48500, 692987, 23200, 37.8, 6.92,15165, 1521, 16985, 2282, 48.45, 2725, 23,  914, 49, 305),
    ("1016x305x487 UB", 62000, 992000, 31700, 40.0, 7.15,19528, 2079, 21871, 3119, 81.81, 5435, 27, 1016, 62, 305),
]

_UC_FULL_EXTENDED = [
    # desig                   A       Iy      Iz    ry     rz   Wely   Welz   Wply   Wplz     Iw      It    tw    h    tf    b
    ("356x406x744 UC",   94800, 360477, 118000, 19.5, 11.16, 14803,  5566, 16580,  8349,  69.9, 22603,  58, 487,  90, 424),
    ("356x406x900 UC",  114600, 458400, 147000, 20.0, 11.32, 18336,  6934, 20536, 10401,  91.9, 42868,  70, 500, 112, 424),
    ("356x406x1086 UC", 138300, 553200, 182000, 20.0, 11.47, 21074,  8585, 23603, 12878, 125.4, 78863,  82, 525, 138, 424),
]


def _bc_section_from_row(row):
    """Create a Section object from a _UB_FULL_EXTENDED or _UC_FULL_EXTENDED row.

    Applies the same unit conversions and alphay/alphaz logic as
    beam_column_table() so the returned object is a drop-in replacement.
    """
    global condition, grade
    desig, A, Iy_cm4, Iz_cm4, ry_cm, rz_cm, Wely, Welz, Wply, Wplz, Iw_dm6, It_cm4, tw, h, tf, b = row
    Iy    = Iy_cm4 * 10000.0        # mm⁴
    Iz    = Iz_cm4 * 10000.0        # mm⁴
    Wpl   = Wply   * 1000.0         # mm³ (major axis plastic)
    Wel   = Wely   * 1000.0         # mm³ (major axis elastic)
    Wpl_z = Wplz   * 1000.0         # mm³ (minor axis plastic)
    Wel_z = Welz   * 1000.0         # mm³ (minor axis elastic)
    Iw    = Iw_dm6 * (100 ** 6)     # mm⁶
    It    = It_cm4 * 10000.0        # mm⁴
    iz    = rz_cm  * 10.0           # mm

    # Imperfection factors – identical logic to beam_column_table()
    if condition == "Rolled":
        if (h / b) <= 1.2 and grade != "S460":
            alphay, alphaz = (0.34, 0.49) if tf <= 100 else (0.76, 0.76)
        elif (h / b) > 1.2 and grade != "S460":
            alphay, alphaz = (0.21, 0.34) if tf <= 40 else (0.34, 0.49)
        else:
            if tf <= 40:
                alphay = alphaz = 0.13
            elif tf > 100:
                alphay = alphaz = 0.49
            else:  # 40 < tf <= 100
                alphay = alphaz = 0.21
    else:
        alphay, alphaz = (0.34, 0.49) if tf <= 40 else (0.49, 0.76)

    sec = Section(desig, A, Iy, Iz, Wpl, Wel, Wpl_z, Wel_z, Iw, It, tw, h, tf, b, iz)
    sec.alphay = alphay
    sec.alphaz = alphaz
    return sec


def _bc_ec3_check(pop, fy, Ned, Mzed, Myed, Lcry, Lcrz, E, G, C1):
    """EC3 §6.3.3 beam-column interaction check for a Section object.

    Returns a result dict when U ≤ 1.0, or None when the section is
    inadequate (or Class 4).  Extracted from beam_column() so the same
    logic can be applied to both xlsx-sourced and extended sections.
    """
    seclass = section_class(fy, pop, Ned)
    pop.seclass = seclass
    if seclass in (1, 2):
        Nrd = pop.A * fy
    elif seclass == 3:
        Nrd = pop.A * fy
    else:
        return None  # Class 4 – skip

    Ncry = (np.pi ** 2 * E * pop.Iy) / (Lcry ** 2)
    Ncrz = (np.pi ** 2 * E * pop.Iz) / (Lcrz ** 2)
    lamy = np.sqrt((pop.A * fy) / Ncry)
    lamz = np.sqrt((pop.A * fy) / Ncrz)
    phiy = 0.5 * (1 + pop.alphay * (lamy - 0.2) + lamy ** 2)
    phiz = 0.5 * (1 + pop.alphaz * (lamz - 0.2) + lamz ** 2)
    chiy = min(1 / (phiy + np.sqrt(phiy ** 2 - lamy ** 2)), 1.0)
    chiz = min(1 / (phiz + np.sqrt(phiz ** 2 - lamz ** 2)), 1.0)
    Nbrd = min(chiy, chiz) * (pop.A * fy)

    LcrLT = Lcry
    Mcr = C1 * (np.pi ** 2 * E * pop.Iz / LcrLT ** 2) * np.sqrt(
        (pop.Iw / pop.Iz) + (LcrLT ** 2 * G * pop.It) / (np.pi ** 2 * E * pop.Iz)
    )

    Wy = pop.Wpl if seclass in (1, 2) else pop.Wel
    lamLT  = np.sqrt((Wy * fy) / Mcr)
    phi_LT = 0.5 * (1 + 0.34 * (lamLT - 0.2) + lamLT ** 2)
    chi_LT = min(1 / (phi_LT + np.sqrt(phi_LT ** 2 - lamLT ** 2)), 1.0)
    Mbrd   = chi_LT * Wy * fy

    Wz   = pop.Wpl_z if seclass in (1, 2) else pop.Wel_z
    Mzrd = Wz * fy

    C_my  = max(0.4, 0.6 + 0.4 * 1.0)  # psi_y = 1.0 (uniform moment)
    C_mz  = C1
    C_mLT = C1

    gamma_M1 = 1.0
    n_y = Ned / (chiy * Nrd / gamma_M1) if chiy > 0 else 0
    n_z = Ned / (chiz * Nrd / gamma_M1) if chiz > 0 else 0

    if seclass in (1, 2):
        k_yy = C_my * min(1 + (lamy - 0.2) * n_y, 1 + 0.8 * n_y)
        k_zz = C_mz * min(1 + (2 * lamz - 0.6) * n_z, 1 + 1.4 * n_z)
        k_yz = 0.6 * k_zz
        if lamy < 0.4:
            k_zy = 0.6 * k_yy
        else:
            k_zy = max(1 - (0.1 * lamz) / max(C_mLT - 0.25, 0.01) * n_z,
                       0.6 * k_yy)
    elif seclass == 3:
        k_yy = C_my * min(1 + 0.6 * lamy * n_y, 1 + 0.6 * n_y)
        k_zz = C_mz * min(1 + 0.6 * lamz * n_z, 1 + 0.6 * n_z)
        k_yz = k_zz
        if lamy < 0.4:
            k_zy = 0.8 * k_zz
        elif lamz < 0.4:
            k_zy = 0.6 + lamz
        else:
            k_zy = max(1 - (0.05 * lamz) / max(C_mLT - 0.25, 0.01) * n_z,
                       0.6 * k_zz)
    else:
        k_yy = 1 + 0.6 * n_y
        k_zz = 1 + 0.6 * n_z
        k_yz = 0.6 * k_zz
        k_zy = 0.6 * k_yy

    k_yy = max(0.1, min(k_yy, 2.0))
    k_zz = max(0.1, min(k_zz, 2.0))
    k_yz = max(0.1, min(k_yz, 2.0))
    k_zy = max(0.1, min(k_zy, 2.0))

    util_y = (Ned / Nbrd) + k_yy * (Myed / Mbrd) + k_yz * (Mzed / Mzrd)
    util_z = (Ned / Nbrd) + k_zy * (Myed / Mbrd) + k_zz * (Mzed / Mzrd)
    U = max(util_y, util_z)

    if U > 1.0:
        return None

    return {
        "Designation": pop.designation,
        "class":       pop.seclass,
        "N_b_Rd":      Nbrd,
        "M_b_Rd":      Mbrd,
        "M_z_Rd":      Mzrd,
        "utilisation": U,
        "chi_y":       chiy,
        "chi_z":       chiz,
        "chi_LT":      chi_LT,
        "k_yy":        k_yy,
        "k_zz":        k_zz,
        "k_yz":        k_yz,
        "k_zy":        k_zy,
        "C_my":        C_my,
        "C_mz":        C_mz,
        "util_y":      util_y,
        "util_z":      util_z,
    }


def table_reader(table, val):
    if table == "CHS":
        ex = openpyxl.load_workbook("CHS.xlsx")
        ex = ex.active
        n = 0.0
        i = 1
        while n<val:
            i = i + 1
            if i > ex.max_row:
                return _chs_extended(val)
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
            if i < 2:
                return _angle_extended(val)
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
            if i > ex.max_row:
                return _uc_extended(val)
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
            if i > ex.max_row:
                return _ub_extended(val)
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
            anet = agross - nh * t * (d + 2.0) + ngs * (((s ** 2) * t) / (4 * p))
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
            anet = agross - nh * t * (d + 2.0) + ngs * (((s ** 2) * t) / (4 * p))
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
            chi = min(1 / (phi + np.sqrt(phi ** 2 - lam_bar ** 2)), 1.0)
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
            chi = min(1 / (phi + np.sqrt(max(phi ** 2 - lam_bar ** 2, 0))), 1.0)
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
            # xlsx exhausted – fall back to extended UB sections
            for ext_row in _UB_FULL_EXTENDED:
                ext_size = ext_row[0]
                Wyy  = ext_row[8]  * 1000.0   # Wpl,y  mm³
                Wzz  = ext_row[9]  * 1000.0   # Wpl,z  mm³
                Wpl  = max(Wyy, Wzz)
                tw   = float(ext_row[12])
                h    = float(ext_row[13])
                tf   = float(ext_row[14])
                hw   = h - 2 * tf
                Av   = hw * tw
                Aw   = hw * tw

                Vpl_Rd = (Av * fy) / (np.sqrt(3) * gamma_M0) / 1000
                if V > Vpl_Rd:
                    continue

                Mpl_Rd = (Wpl * fy) / gamma_M0 / 1e6
                if V / Vpl_Rd <= 0.5:
                    M_Rd = Mpl_Rd
                else:
                    rho  = (2 * V / Vpl_Rd - 1) ** 2
                    M_Rd = ((Wpl - (rho * Aw ** 2) / (4 * tw)) * fy / gamma_M0) / 1e6

                if M <= M_Rd:
                    return {
                        "Type": "Restrained Beam",
                        "Size": ext_size,
                        "M_Rd (kNm)": round(M_Rd, 2),
                        "V_Rd (kN)": round(Vpl_Rd, 2),
                        "Utilization (%)": round((M / M_Rd) * 100, 2),
                    }

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
            # xlsx exhausted – fall back to extended UB sections
            for ext_row in _UB_FULL_EXTENDED:
                ext_size = ext_row[0]
                Wpl  = ext_row[8]  * 1000.0    # Wpl,y  mm³
                iz   = ext_row[5]  * 10.0       # rz (cm) → iz (mm)
                tw   = float(ext_row[12])
                h    = float(ext_row[13])
                tf   = float(ext_row[14])
                b    = float(ext_row[15])

                if endcondition == "Free":
                    k = 1.0
                elif endcondition == "Partial":
                    k = 0.85
                elif endcondition == "Full":
                    k = 0.7
                elif endcondition == "Cantilever":
                    k = 2.0
                else:
                    raise ValueError("Endcondition not recognized.")

                lamz    = (k * L) / iz
                laml    = np.pi * np.sqrt(E / fy)
                lamzba  = lamz / laml
                lamltb  = 0.9 * lamzba  # EC3 §6.3.2.2: C1=1.0, u=0.9, v=1.0, βw=1.0
                hoverb  = h / b

                if condition == "Rolled":
                    alt = 0.34 if hoverb <= 2 else 0.49
                    phi = 0.5 * (1 + alt * (lamltb - 0.4) + 0.75 * lamltb ** 2)
                    chi = min(1 / (phi + np.sqrt(phi ** 2 - 0.75 * lamltb ** 2)), 1.0)
                else:
                    if hoverb <= 2 and condition == "Welded":
                        alt = 0.49
                    elif hoverb > 2 and condition == "Welded":
                        alt = 0.76
                    else:
                        alt = 0.76
                    phi = 0.5 * (1 + alt * (lamltb - 0.2) + lamltb ** 2)
                    chi = min(1 / (phi + np.sqrt(phi ** 2 - lamltb ** 2)), 1.0)

                hw     = h - 2 * tf
                Av     = hw * tw
                Aw     = hw * tw
                Vpl_Rd = (Av * fy) / (np.sqrt(3) * gamma_M1) / 1000

                if V > Vpl_Rd:
                    continue

                if V > 0.5 * Vpl_Rd:
                    rho       = ((2 * V / Vpl_Rd) - 1) ** 2
                    W_reduced = Wpl - (rho * Aw ** 2) / (4 * tw)
                else:
                    W_reduced = Wpl

                Mbrd = (chi * W_reduced * fy / gamma_M1) / 1e6

                if M <= Mbrd:
                    return {
                        "Type": "Unrestrained Beam",
                        "Size": ext_size,
                        "x_LT": round(chi, 3),
                        "Mb_Rd (kNm)": round(Mbrd, 2),
                        "Vpl_Rd": round(Vpl_Rd, 2),
                        "Utilization (%)": round((M / Mbrd) * 100, 2),
                    }

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
        # EC3 Table 5.2 — Class 3 web limit for the compression-dominant range (α > 0.5).
        # ψ_EC3 is the stress ratio at the web faces (tension-positive convention):
        #   ψ_EC3 = 1 − 1/α  maps α ∈ (0.5, 1.0) → ψ_EC3 ∈ (−1, 0)
        # The denominator 0.67 + 0.33·ψ_EC3 = 1 − 0.33/α is bounded in [0.34, 0.67]
        # for α ∈ (0.5, 1), so division by zero cannot occur.
        # α = 1 (full compression, ψ_EC3 = 1) is treated separately as 42ε.
        if alpha >= 1.0:
            limit3 = 42.0 * eps                                 # EC3: ψ=1, uniform compression
        else:
            # 0.5 < alpha < 1.0 — neutral axis is within the web
            psi_EC3 = 1.0 - 1.0 / alpha                        # ψ_EC3 ∈ (−1, 0)
            ec3_limit3 = 42.0 * eps / (0.67 + 0.33 * psi_EC3)  # 124ε at α→0.5, 62.7ε at α→1
            limit3 = max(limit2, ec3_limit3)                    # ensure Class 3 ≥ Class 2

    else:
        # Prevent division by zero
        if alpha <= 0:
            return 4

        limit1 = 36.0 * eps / alpha
        limit2 = 41.5 * eps / alpha
        # EC3 Table 5.2 — Class 3 web limit for tension-dominant case (α ≤ 0.5).
        # At α = 0.5 (pure bending, ψ_EC3 = −1): EC3 gives 124ε.
        # For α < 0.5 the compression zone is smaller, so the limit is even more
        # permissive; 124ε is a conservative (safe) lower bound.
        limit3 = max(limit2, 124.0 * eps)

    if lambda_w <= limit1:
        web_class = 1
    elif lambda_w <= limit2:
        web_class = 2
    elif lambda_w <= limit3:
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
    def __init__(self, designation, A, Iy, Iz, Wpl, Wel, Wpl_z, Wel_z, Iw, It, tw, h, tf, b, iz, seclass=None, alphay=None, alphaz=None):
        self.designation = designation
        self.A = A
        self.Iy = Iy
        self.Iz = Iz
        self.Wpl = Wpl       # major-axis plastic modulus W_pl,y
        self.Wel = Wel       # major-axis elastic modulus W_el,y
        self.Wpl_z = Wpl_z  # minor-axis plastic modulus W_pl,z
        self.Wel_z = Wel_z  # minor-axis elastic modulus W_el,z
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
    Wpl = float(ex.cell(row=i, column=8).value) * 1000    # W_pl,y  (major axis, cm³ → mm³)
    Wel = float(ex.cell(row=i, column=6).value) * 1000    # W_el,y  (major axis, cm³ → mm³)
    Wpl_z = float(ex.cell(row=i, column=9).value) * 1000  # W_pl,z  (minor axis, cm³ → mm³)
    Wel_z = float(ex.cell(row=i, column=7).value) * 1000  # W_el,z  (minor axis, cm³ → mm³)
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



    outer = Section(designation, A, Iy, Iz, Wpl, Wel, Wpl_z, Wel_z, Iw, It, tw, h, tf, b, iz)
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
    A0 = Ned / fy
    if shape == "UB":
        ex = openpyxl.load_workbook("UB-2.xlsx").active
        ext_list = _UB_FULL_EXTENDED
    elif shape == "UC":
        ex = openpyxl.load_workbook("UC-2.xlsx").active
        ext_list = _UC_FULL_EXTENDED
    else:
        raise ValueError("Unknown shape")

    # --- Phase 1: iterate through xlsx sections ---
    for i in range(2, ex.max_row + 1):
        cell_val = ex.cell(row=i, column=14).value
        if cell_val is None:
            break
        if not float(cell_val) > A0:
            continue
        pop    = beam_column_table(shape, i)
        result = _bc_ec3_check(pop, fy, Ned, Mzed, Myed, Lcry, Lcrz, E, G, C1)
        if result is not None:
            return result

    # --- Phase 2: extended section fallback ---
    for row in ext_list:
        if not float(row[1]) > A0:
            continue
        pop    = _bc_section_from_row(row)
        result = _bc_ec3_check(pop, fy, Ned, Mzed, Myed, Lcry, Lcrz, E, G, C1)
        if result is not None:
            return result

    raise ValueError("No suitable section found.")




