"""
frame_analysis.py
2-D rigid frame analysis using the direct stiffness method.

Conventions
-----------
* Units: metres (m), kilonewtons (kN), kilonewton-metres (kNm).
* Each node has 3 DOFs: u (horizontal +ve right), v (vertical +ve up), θ (rotation +ve CCW).
* Member local x-axis runs from start-node to end-node.
* Positive axial force  → tension.
* Positive shear force  → local +y direction at the start-node face.
* Positive bending moment → sagging (local +y concave).
"""

import numpy as np

# ─────────────────────────────────────────────────────────────────────────────
# Representative section properties used for the elastic analysis
# (actual design sections are chosen afterwards)
# Based on mid-range steel I-sections:
#   Beam  → approx. UB 457×191×67 : I_y = 29 400 cm⁴, A = 85.4 cm²
#   Column → approx. UC 254×254×73 : I_y = 11 360 cm⁴, A = 93.1 cm²
# ─────────────────────────────────────────────────────────────────────────────
_E        = 210e6          # kN/m²  (210 GPa)
_I_BEAM   = 2.94e-4        # m⁴
_A_BEAM   = 8.54e-3        # m²
_I_COL    = 1.136e-4       # m⁴
_A_COL    = 9.31e-3        # m²


# ─────────────────────────────────────────────────────────────────────────────
# Internal helpers
# ─────────────────────────────────────────────────────────────────────────────

def _local_stiffness(E, A, I, L):
    """Return the 6×6 Euler-Bernoulli beam-column local stiffness matrix."""
    EA_L   = E * A / L
    EI_L3  = E * I / L ** 3
    EI_L2  = E * I / L ** 2
    EI_L   = E * I / L
    return np.array([
        [ EA_L,        0,          0,       -EA_L,       0,          0      ],
        [ 0,      12*EI_L3,   6*EI_L2,       0,     -12*EI_L3,   6*EI_L2  ],
        [ 0,       6*EI_L2,   4*EI_L,        0,      -6*EI_L2,   2*EI_L   ],
        [-EA_L,        0,          0,        EA_L,       0,          0      ],
        [ 0,     -12*EI_L3,  -6*EI_L2,       0,      12*EI_L3,  -6*EI_L2  ],
        [ 0,       6*EI_L2,   2*EI_L,        0,      -6*EI_L2,   4*EI_L   ],
    ])


def _transform(angle):
    """Return the 6×6 transformation matrix (local → global) for member angle α."""
    c, s = np.cos(angle), np.sin(angle)
    return np.array([
        [ c,  s,  0,  0,  0,  0],
        [-s,  c,  0,  0,  0,  0],
        [ 0,  0,  1,  0,  0,  0],
        [ 0,  0,  0,  c,  s,  0],
        [ 0,  0,  0, -s,  c,  0],
        [ 0,  0,  0,  0,  0,  1],
    ])


def _fef_local(p_x, p_y, L):
    """
    Fixed-end forces in local coordinates for a uniform distributed load.
    p_x : load per unit length along member (+ve in local x-direction)  [kN/m]
    p_y : load per unit length perp. to member (+ve in local y-direction) [kN/m]
    Returns 6-element vector [F1x, F1y, M1, F2x, F2y, M2].
    """
    return np.array([
        p_x * L / 2,
        p_y * L / 2,
        p_y * L ** 2 / 12,
        p_x * L / 2,
        p_y * L / 2,
       -p_y * L ** 2 / 12,
    ])


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def analyse_frame(nodes, members, supports, node_loads, udl_loads):
    """
    Perform a linear-elastic 2-D rigid frame analysis.

    Parameters
    ----------
    nodes : list of [node_id (int), x (float, m), y (float, m)]
    members : list of [member_id (int), start_node (int), end_node (int),
                       member_type (str)]
        member_type must be "Beam" or "Column".
    supports : list of [node_id (int), support_type (str)]
        support_type: "Fixed" | "Pinned" | "Roller (H)" | "Roller (V)"
    node_loads : list of [node_id (int), Fx (kN), Fy (kN), Mz (kNm)]
    udl_loads  : list of [member_id (int), wx (kN/m), wy (kN/m)]
        wx: horizontal component of distributed load (+ve = rightward).
        wy: vertical component of distributed load   (+ve = upward).

    Returns
    -------
    member_results : dict  {member_id → dict with keys:
        "N_start"  – axial force at start node (kN, +ve = tension)
        "V_start"  – shear force at start node (kN)
        "M_start"  – bending moment at start node (kNm)
        "N_end"    – axial force at end node   (kN, +ve = tension)
        "V_end"    – shear force at end node   (kN)
        "M_end"    – bending moment at end node (kNm)
        "length"   – member length (m)
        "type"     – "Beam" or "Column"
        "angle_deg"– inclination from horizontal (degrees)
    }
    node_displacements : dict  {node_id → {"u": float, "v": float, "theta": float}}
    """

    # ── index maps ────────────────────────────────────────────────────────
    node_map = {int(n[0]): i for i, n in enumerate(nodes)}
    n_nodes  = len(nodes)
    n_dof    = 3 * n_nodes

    # ── global arrays ─────────────────────────────────────────────────────
    K = np.zeros((n_dof, n_dof))
    F = np.zeros(n_dof)

    # pre-process UDL and support lookups
    udl_map     = {int(u[0]): (float(u[1]), float(u[2])) for u in udl_loads}
    support_map = {int(s[0]): s[1] for s in supports}

    member_meta = {}   # member_id → geometry + stiffness data

    # ── Assemble stiffness matrix & load vector ────────────────────────────
    for mem in members:
        mid, ni, nj, mtype = int(mem[0]), int(mem[1]), int(mem[2]), str(mem[3])

        xi, yi = float(nodes[node_map[ni]][1]), float(nodes[node_map[ni]][2])
        xj, yj = float(nodes[node_map[nj]][1]), float(nodes[node_map[nj]][2])

        L     = float(np.sqrt((xj - xi) ** 2 + (yj - yi) ** 2))
        if L < 1e-12:
            raise ValueError(f"Member {mid} has zero length. Check node coordinates.")
        angle = float(np.arctan2(yj - yi, xj - xi))

        I = _I_BEAM if mtype == "Beam" else _I_COL
        A = _A_BEAM if mtype == "Beam" else _A_COL

        k_loc  = _local_stiffness(_E, A, I, L)
        T      = _transform(angle)
        K_glob = T.T @ k_loc @ T

        i_idx = node_map[ni]
        j_idx = node_map[nj]
        dofs  = [3*i_idx, 3*i_idx+1, 3*i_idx+2,
                 3*j_idx, 3*j_idx+1, 3*j_idx+2]

        for a in range(6):
            for b in range(6):
                K[dofs[a], dofs[b]] += K_glob[a, b]

        # Fixed-end forces from UDL (if any)
        fef_loc = np.zeros(6)
        if mid in udl_map:
            wx_g, wy_g = udl_map[mid]
            c, s = np.cos(angle), np.sin(angle)
            # Decompose global UDL into local coords
            p_x =  wx_g * c + wy_g * s
            p_y = -wx_g * s + wy_g * c
            fef_loc = _fef_local(p_x, p_y, L)
            fef_glob = T.T @ fef_loc
            for a in range(6):
                F[dofs[a]] += fef_glob[a]

        member_meta[mid] = {
            "length": L,
            "angle":  angle,
            "k_loc":  k_loc,
            "T":      T,
            "dofs":   dofs,
            "type":   mtype,
            "fef_loc": fef_loc,
        }

    # ── Nodal loads ───────────────────────────────────────────────────────
    for load in node_loads:
        nid, Fx, Fy, Mz = int(load[0]), float(load[1]), float(load[2]), float(load[3])
        if nid not in node_map:
            raise ValueError(f"Node load references unknown node {nid}.")
        idx = node_map[nid]
        F[3*idx]   += Fx
        F[3*idx+1] += Fy
        F[3*idx+2] += Mz

    # ── Boundary conditions ───────────────────────────────────────────────
    constrained = set()
    for nid, stype in support_map.items():
        if nid not in node_map:
            raise ValueError(f"Support references unknown node {nid}.")
        idx = node_map[nid]
        stype_lower = str(stype).strip().lower()
        if stype_lower == "fixed":
            constrained.update([3*idx, 3*idx+1, 3*idx+2])
        elif stype_lower == "pinned":
            constrained.update([3*idx, 3*idx+1])
        elif stype_lower in ("roller (h)", "roller_h", "roller-h"):
            constrained.add(3*idx)
        elif stype_lower in ("roller (v)", "roller_v", "roller-v"):
            constrained.add(3*idx+1)
        else:
            raise ValueError(f"Unknown support type '{stype}' at node {nid}.")

    free_dofs = [d for d in range(n_dof) if d not in constrained]
    if len(free_dofs) == 0:
        raise ValueError("No free DOFs — all nodes appear to be fully restrained.")

    # ── Solve ─────────────────────────────────────────────────────────────
    K_free = K[np.ix_(free_dofs, free_dofs)]
    F_free = F[free_dofs]

    try:
        D_free = np.linalg.solve(K_free, F_free)
    except np.linalg.LinAlgError:
        raise ValueError(
            "Singular stiffness matrix — the frame may be a mechanism. "
            "Check that all members are connected and supports are adequate."
        )

    D = np.zeros(n_dof)
    for i, dof in enumerate(free_dofs):
        D[dof] = D_free[i]

    # ── Node displacements ────────────────────────────────────────────────
    node_displacements = {}
    for nid, idx in node_map.items():
        node_displacements[nid] = {
            "u":     D[3*idx],
            "v":     D[3*idx+1],
            "theta": D[3*idx+2],
        }

    # ── Member force recovery ─────────────────────────────────────────────
    member_results = {}
    for mid, meta in member_meta.items():
        dofs    = meta["dofs"]
        T       = meta["T"]
        k_loc   = meta["k_loc"]
        fef_loc = meta["fef_loc"]
        L       = meta["length"]
        angle   = meta["angle"]
        mtype   = meta["type"]

        d_glob  = D[dofs]
        d_loc   = T @ d_glob
        f_loc   = k_loc @ d_loc - fef_loc

        # f_loc = [F1x, F1y, M1, F2x, F2y, M2]
        # Positive axial (F1x > 0) means the member pushes the start-node in +local-x
        # → the member itself is in compression.  We adopt tension-positive internally.
        member_results[mid] = {
            "N_start":   -f_loc[0],       # tension positive
            "V_start":    f_loc[1],
            "M_start":    f_loc[2],
            "N_end":      f_loc[3],        # tension positive at end face
            "V_end":     -f_loc[4],
            "M_end":      f_loc[5],
            "length":     L,
            "type":       mtype,
            "angle_deg":  float(np.degrees(angle)),
        }

    return member_results, node_displacements
