"""Independent Python reference implementation of the P3 signal model (used for validation)."""
import math

def hcm_delay(C, g, X, cap, T=0.25, k=0.5, I=1.0):
    """HCM 6th edition control delay components for one lane group (s/veh)."""
    lam = g / C
    d1 = 0.5 * C * (1 - lam) ** 2 / (1 - min(1.0, X) * lam)
    d2 = 900 * T * ((X - 1) + math.sqrt((X - 1) ** 2 + 8 * k * I * X / (cap * T)))
    return d1, d2

def webster_delay(C, g, X, q_vph):
    lam = g / C; q = q_vph / 3600
    if X >= 1: return float("nan")
    return C * (1 - lam) ** 2 / (2 * (1 - lam * X)) + X ** 2 / (2 * q * (1 - X)) - 0.65 * (C / q ** 2) ** (1 / 3) * X ** (2 + 5 * lam)

def node_model(Cb, gm, gs, clr, X_main, v_main, v_side, lanes_main=2, lanes_side=1, growth=0.0, cmin=60, cmax=150, rnd=5, adopted=None):
    L = 2 * clr
    ge_m = (Cb - L) * gm / (gm + gs); ge_s = (Cb - L) * gs / (gm + gs)
    y_m0 = X_main * ge_m / Cb
    y_s0 = y_m0 * (v_side / lanes_side) / (v_main / lanes_main)
    s_m = v_main / y_m0; s_s = v_side / y_s0             # implied saturation flows (veh/h)
    f = 1 + growth
    vm, vs = v_main * f, v_side * f
    y_m, y_s = y_m0 * f, y_s0 * f
    Y = y_m + y_s
    # before
    Xb_m, Xb_s = y_m * Cb / ge_m, y_s * Cb / ge_s
    b_m = hcm_delay(Cb, ge_m, Xb_m, s_m * ge_m / Cb); b_s = hcm_delay(Cb, ge_s, Xb_s, s_s * ge_s / Cb)
    d_before = (vm * sum(b_m) + vs * sum(b_s)) / (vm + vs)
    # after
    C0 = (1.5 * L + 5) / (1 - Y) if Y < 0.95 else cmax
    C0 = min(cmax, max(cmin, C0))
    Ca = adopted if adopted else math.ceil(C0 / rnd) * rnd
    ga_m = (Ca - L) * y_m / Y; ga_s = (Ca - L) * y_s / Y
    Xa_m, Xa_s = y_m * Ca / ga_m, y_s * Ca / ga_s
    a_m = hcm_delay(Ca, ga_m, Xa_m, s_m * ga_m / Ca); a_s = hcm_delay(Ca, ga_s, Xa_s, s_s * ga_s / Ca)
    d_after = (vm * sum(a_m) + vs * sum(a_s)) / (vm + vs)
    return dict(L=L, Y=Y, C0=C0, Ca=Ca, d_before=d_before, d_after=d_after, red=1 - d_after / d_before,
                Xb_m=Xb_m, Xb_s=Xb_s, Xa=Xa_m, b_m=b_m, b_s=b_s, a_m=a_m, a_s=a_s, ge_m=ge_m, ge_s=ge_s, ga_m=ga_m, ga_s=ga_s, s_m=s_m, s_s=s_s)
