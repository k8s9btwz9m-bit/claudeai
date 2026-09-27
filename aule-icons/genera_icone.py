"""Genera le icone SVG di Aulë. Riusa palette e helper delle icone Arkenstone."""
import math, os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "arkenstone-icons"))
from genera_icone import f, pts, shade, frame, sparkle, G, B_PALE, B_LIGHT, B_MID, BG2

OUT = os.path.dirname(os.path.abspath(__file__))

ACCENT = f'''    <linearGradient id="accent" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="{B_PALE}"/><stop offset="1" stop-color="{B_MID}"/>
    </linearGradient>
    <linearGradient id="steel" x1="0" y1="0" x2="0.3" y2="1">
      <stop offset="0" stop-color="{G[2]}"/><stop offset="1" stop-color="{G[4]}"/>
    </linearGradient>
    <linearGradient id="table" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#FFFFFF"/><stop offset="1" stop-color="{G[1]}"/>
    </linearGradient>
'''

def gem(cx, cy, R, uid, grid=True, lw=None):
    """La gemma dell'Arkenstone (taglio ottagonale) centrata in cx,cy con raggio R."""
    r = R * 0.54
    lw = lw if lw is not None else max(1.2, R / 140)
    ang = [math.radians(22.5 + 45*i - 90) for i in range(8)]
    O = [(cx + R*math.cos(a), cy + R*math.sin(a)) for a in ang]
    I = [(cx + r*math.cos(a), cy + r*math.sin(a)) for a in ang]
    s = ""
    for i in range(8):
        j = (i + 1) % 8
        M = ((O[i][0]+O[j][0])/2, (O[i][1]+O[j][1])/2)
        am = (ang[i] + ang[j]) / 2
        a1, a2 = ang[i] + math.radians(11), ang[j] - math.radians(11)
        s += f'    <polygon points="{pts([I[i], I[j], M])}" fill="{shade(math.cos(am)*0.55, math.sin(am)*0.55, 0.85)}"/>\n'
        s += f'    <polygon points="{pts([O[i], M, I[i]])}" fill="{shade(math.cos(a1), math.sin(a1), 0.45)}"/>\n'
        s += f'    <polygon points="{pts([M, O[j], I[j]])}" fill="{shade(math.cos(a2), math.sin(a2), 0.45)}"/>\n'
    M = ((O[1][0]+O[2][0])/2, (O[1][1]+O[2][1])/2)
    s += f'    <polygon points="{pts([I[1], I[2], M])}" fill="url(#accent)"/>\n'
    s += f'    <polygon points="{pts([M, O[2], I[2]])}" fill="{B_LIGHT}" opacity="0.55"/>\n'
    d = ""
    for i in range(8):
        j = (i + 1) % 8
        M = ((O[i][0]+O[j][0])/2, (O[i][1]+O[j][1])/2)
        d += f"M{f(O[i][0])},{f(O[i][1])} L{f(I[i][0])},{f(I[i][1])} M{f(M[0])},{f(M[1])} L{f(I[i][0])},{f(I[i][1])} M{f(M[0])},{f(M[1])} L{f(I[j][0])},{f(I[j][1])} "
    s += f'    <path d="{d}" stroke="#FFFFFF" stroke-opacity="0.55" stroke-width="{f(lw)}" fill="none"/>\n'
    s += f'    <polygon points="{pts(O)}" fill="none" stroke="{G[5]}" stroke-width="{f(lw*2)}" stroke-linejoin="round"/>\n'
    s += f'    <polygon points="{pts(I)}" fill="url(#table)"/>\n'
    if grid:
        c = r / 3 * 1.1
        s += f'    <clipPath id="tc{uid}"><polygon points="{pts(I)}"/></clipPath>\n'
        s += f'    <rect x="{f(cx + c + lw*2)}" y="{f(cy - r)}" width="{f(r)}" height="{f(r - c - lw*2)}" fill="{B_LIGHT}" opacity="0.6" clip-path="url(#tc{uid})"/>\n'
        g = "".join(f"M{f(cx + k*c)},{f(cy - r)} V{f(cy + r)} M{f(cx - r)},{f(cy + k*c)} H{f(cx + r)} " for k in (-1, 1))
        s += f'    <path d="{g}" stroke="{G[3]}" stroke-width="{f(lw*1.5)}" clip-path="url(#tc{uid})"/>\n'
    s += f'    <polygon points="{pts(I)}" fill="none" stroke="#FFFFFF" stroke-width="{f(lw*1.5)}" stroke-opacity="0.9" stroke-linejoin="round"/>\n'
    return s

def facet_poly(poly, nx, ny, nz=0.4):
    return f'    <polygon points="{pts(poly)}" fill="{shade(nx, ny, nz)}"/>\n'

def rot(p, a, c=(0, 0)):
    x, y = p[0]-c[0], p[1]-c[1]
    return (c[0] + x*math.cos(a) - y*math.sin(a), c[1] + x*math.sin(a) + y*math.cos(a))

# ---------------- 1. Incudine con la Gemma ----------------
def icon1():
    s = f'    <circle cx="512" cy="470" r="400" fill="url(#halo)"/>\n'
    s += f'    <ellipse cx="512" cy="822" rx="300" ry="18" fill="#000" opacity="0.3"/>\n'
    T = 520   # quota del piano
    # sagoma a segmenti dritti (stile nanico)
    top      = [(318,T), (806,T), (792,T+22), (330,T+22)]
    horn_up  = [(318,T), (330,T+22), (200,T+40)]
    horn_dn  = [(200,T+40), (330,T+22), (392,T+62), (392,T+92)]
    face     = [(330,T+22), (792,T+22), (792,T+92), (392,T+92), (392,T+62)]
    heel     = [(792,T+22), (806,T), (806,T+76), (792,T+92)]
    under    = [(392,T+92), (792,T+92), (716,T+122), (440,T+122)]
    waist    = [(440,T+122), (716,T+122), (670,T+200), (486,T+200)]
    foot_top = [(486,T+200), (670,T+200), (760,T+250), (396,T+250)]
    foot     = [(396,T+250), (760,T+250), (760,T+290), (396,T+290)]
    s += f'    <polygon points="{pts(face)}" fill="url(#steel)"/>\n'
    s += f'    <polygon points="{pts(horn_dn)}" fill="{G[4]}"/>\n'
    s += facet_poly(horn_up, -0.6, -0.7, 0.4)
    s += facet_poly(top, 0, -1, 0.9)
    s += facet_poly(heel, 1, 0, 0.3)
    s += facet_poly(under, 0.2, 0.9, 0.3)
    s += f'    <polygon points="{pts(waist)}" fill="{G[4]}"/>\n'
    s += f'    <polygon points="{pts(foot_top)}" fill="{G[2]}"/>\n'
    s += f'    <polygon points="{pts(foot)}" fill="{G[3]}"/>\n'
    # griglia incisa sulla vita
    gx0, gx1, gy0, gy1 = 505, 651, T+138, T+188
    g = f"M{gx0},{gy0} H{gx1} M{gx0},{gy1} H{gx1} M{gx0},{gy0} V{gy1} M{gx1},{gy0} V{gy1} "
    for k in (1, 2):
        g += f"M{f(gx0 + (gx1-gx0)*k/3)},{gy0} V{gy1} "
    g += f"M{gx0},{f((gy0+gy1)/2)} H{gx1} "
    s += f'    <path d="{g}" stroke="{G[5]}" stroke-width="3" fill="none"/>\n'
    s += f'    <rect x="{f(gx0 + (gx1-gx0)*2/3 + 3)}" y="{gy0 + 3}" width="{f((gx1-gx0)/3 - 6)}" height="{f((gy1-gy0)/2 - 6)}" fill="{B_LIGHT}" opacity="0.8"/>\n'
    # contorno
    outline = [(318,T),(806,T),(806,T+76),(792,T+92),(716,T+122),(670,T+200),(760,T+250),(760,T+290),(396,T+290),(396,T+250),(486,T+200),(440,T+122),(392,T+92),(392,T+62),(200,T+40)]
    s += f'    <polygon points="{pts(outline)}" fill="none" stroke="#1F2329" stroke-width="4" stroke-linejoin="round"/>\n'
    s += f'    <path d="M318,{T} H806" stroke="#FFFFFF" stroke-opacity="0.7" stroke-width="3"/>\n'
    # gemma sul piano
    R = 150; cy = T - R*math.cos(math.radians(22.5)) - 2
    s += f'    <ellipse cx="560" cy="{T+2}" rx="120" ry="8" fill="{B_LIGHT}" opacity="0.35"/>\n'
    s += gem(560, cy, R, "a")
    s += sparkle(560 - R*0.42, cy - R*0.93, 30, "#FFFFFF")
    return frame(s, ACCENT)

# ---------------- 3. Martello e Scintilla ----------------
def icon3():
    cx, cy = 588, 580
    s = f'    <circle cx="{cx}" cy="{cy}" r="360" fill="url(#halo)"/>\n'
    s += '    <radialGradient id="core" cx="50%" cy="50%" r="50%"><stop offset="0" stop-color="#FFFFFF"/><stop offset="0.45" stop-color="' + B_PALE + '"/><stop offset="1" stop-color="' + B_LIGHT + '" stop-opacity="0"/></radialGradient>\n'
    # 8 raggi allineati ai vertici dell'ottagono, lunghezze diverse (grafico radiale)
    lens = [300, 190, 250, 160, 280, 175, 230, 150]
    for k in range(8):
        a = math.radians(22.5 + 45*k - 90)
        L = lens[k]; w = 16 if L > 200 else 11
        tip = (cx + L*math.cos(a), cy + L*math.sin(a))
        base = (cx + 30*math.cos(a), cy + 30*math.sin(a))
        px, py = -math.sin(a)*w, math.cos(a)*w
        mid = (cx + 70*math.cos(a), cy + 70*math.sin(a))
        col = "#FFFFFF" if L > 200 else B_LIGHT
        s += f'    <polygon points="{pts([base, (mid[0]+px, mid[1]+py), tip, (mid[0]-px, mid[1]-py)])}" fill="{col}"/>\n'
    # piccoli frammenti
    for a, d, r in [(200, 215, 7), (-20, 205, 6), (120, 150, 5), (60, 250, 5), (-100, 225, 6), (160, 280, 4)]:
        x, y = cx + d*math.cos(math.radians(a)), cy + d*math.sin(math.radians(a))
        s += f'    <rect x="{f(x-r)}" y="{f(y-r)}" width="{2*r}" height="{2*r}" transform="rotate(45 {f(x)} {f(y)})" fill="{B_PALE}"/>\n'
    # martello (coordinate locali: faccia di battuta a x=0, manico verso l'alto)
    head = [(-250, -78), (-20, -78), (0, -58), (0, 58), (-20, 78), (-250, 78), (-270, 58), (-270, -58)]
    top  = [(-250, -78), (-20, -78), (-38, -56), (-232, -56)]
    faceR = [(-20, -78), (0, -58), (0, 58), (-20, 78), (-38, 56), (-38, -56)]
    bot  = [(-232, 56), (-38, 56), (-20, 78), (-250, 78)]
    faceL = [(-270, -58), (-250, -78), (-232, -56), (-232, 56), (-250, 78), (-270, 58)]
    band = [(-150, -80), (-120, -80), (-120, 80), (-150, 80)]
    handle = [(-157, -78), (-113, -78), (-108, -440), (-162, -440)]
    a = math.radians(28); o = (cx - 34, cy)
    T = lambda poly: [(o[0] + p[0]*math.cos(a) - p[1]*math.sin(a), o[1] + p[0]*math.sin(a) + p[1]*math.cos(a)) for p in poly]
    s += f'    <polygon points="{pts(T(handle))}" fill="{G[4]}" stroke="#1F2329" stroke-width="4" stroke-linejoin="round"/>\n'
    for yy in range(-400, -120, 46):
        s += f'    <polygon points="{pts(T([(-162, yy), (-108, yy - 14), (-108, yy + 4), (-162, yy + 18)]))}" fill="{G[5]}"/>\n'
    s += f'    <polygon points="{pts(T(head))}" fill="url(#steel)"/>\n'
    s += f'    <polygon points="{pts(T(top))}" fill="{G[1]}"/>\n'
    s += f'    <polygon points="{pts(T(faceR))}" fill="{G[2]}"/>\n'
    s += f'    <polygon points="{pts(T(bot))}" fill="{G[5]}"/>\n'
    s += f'    <polygon points="{pts(T(faceL))}" fill="{G[3]}"/>\n'
    s += f'    <polygon points="{pts(T(band))}" fill="{G[4]}"/>\n'
    s += f'    <polygon points="{pts(T(head))}" fill="none" stroke="#1F2329" stroke-width="4" stroke-linejoin="round"/>\n'
    s += f'    <polyline points="{pts(T([(-20, -76), (-2, -58), (-2, 58)]))}" fill="none" stroke="#FFFFFF" stroke-opacity="0.8" stroke-width="3"/>\n'
    # nucleo della scintilla davanti alla faccia del martello
    s += f'    <circle cx="{cx}" cy="{cy}" r="95" fill="url(#core)"/>\n'
    s += sparkle(cx, cy, 52, "#FFFFFF")
    return frame(s, ACCENT)

# ---------------- 5. La Fucina ----------------
def icon5():
    s = ""
    # muro di blocchi sfalsati
    tones = [G[4], G[5], "#6B7380", G[5], "#747C89"]
    rows = [(300 + 88*i) for i in range(-4, 7)]
    k = 0
    for ri, y in enumerate(rows):
        off = -100 if ri % 2 else 0
        x = off - 20
        while x < 1044:
            w = 196 if (k % 3) else 164
            s += f'    <rect x="{x+4}" y="{y+4}" width="{w-8}" height="80" rx="6" fill="{tones[k % 5]}" opacity="0.55"/>\n'
            x += w; k += 1
    s += f'    <rect width="1024" height="1024" fill="url(#vignette)"/>\n'
    # apertura: porta nanica (lati verticali + timpano trapezoidale)
    door = [(352, 830), (352, 480), (430, 370), (594, 370), (672, 480), (672, 830)]
    frame_out = [(300, 830), (300, 462), (400, 318), (624, 318), (724, 462), (724, 830)]
    s += f'    <polygon points="{pts(frame_out)}" fill="{G[3]}" stroke="#1F2329" stroke-width="4" stroke-linejoin="round"/>\n'
    # conci dello stipite
    for a, b in [((300, 462), (352, 480)), ((400, 318), (430, 370)), ((624, 318), (594, 370)), ((724, 462), (672, 480)),
                 ((300, 600), (352, 600)), ((300, 720), (352, 720)), ((724, 600), (672, 600)), ((724, 720), (672, 720)),
                 ((512, 318), (512, 370))]:
        s += f'    <line x1="{a[0]}" y1="{a[1]}" x2="{b[0]}" y2="{b[1]}" stroke="#1F2329" stroke-width="4"/>\n'
    s += f'    <polyline points="{pts(frame_out[:4])}" fill="none" stroke="#FFFFFF" stroke-opacity="0.45" stroke-width="3"/>\n'
    s += f'    <polygon points="{pts([(400,318),(624,318),(594,370),(430,370)])}" fill="{G[2]}"/>\n'
    s += f'    <polygon points="{pts(door)}" fill="url(#forge)"/>\n'
    s += f'    <polygon points="{pts(door)}" fill="none" stroke="#1F2329" stroke-width="5" stroke-linejoin="round"/>\n'
    # bagliore sul pavimento
    s += f'    <rect x="0" y="830" width="1024" height="194" fill="#15181C"/>\n'
    s += f'    <ellipse cx="512" cy="836" rx="300" ry="36" fill="url(#spill)"/>\n'
    s += f'    <line x1="0" y1="830" x2="1024" y2="830" stroke="{G[5]}" stroke-width="4"/>\n'
    s += gem(512, 620, 104, "f")
    s += sparkle(470, 520, 22, "#FFFFFF")
    defs = ACCENT + f'''    <radialGradient id="forge" cx="50%" cy="60%" r="65%">
      <stop offset="0" stop-color="#FFFFFF"/><stop offset="0.35" stop-color="{B_PALE}"/>
      <stop offset="0.75" stop-color="{B_LIGHT}"/><stop offset="1" stop-color="{B_MID}"/>
    </radialGradient>
    <radialGradient id="spill" cx="50%" cy="50%" r="50%">
      <stop offset="0" stop-color="{B_LIGHT}" stop-opacity="0.7"/><stop offset="1" stop-color="{B_LIGHT}" stop-opacity="0"/>
    </radialGradient>
    <radialGradient id="vignette" cx="50%" cy="58%" r="70%">
      <stop offset="0.3" stop-color="{B_LIGHT}" stop-opacity="0.12"/><stop offset="1" stop-color="{BG2}" stop-opacity="0.85"/>
    </radialGradient>
'''
    return frame(s, defs)

# ---------------- 6. Il Sigillo di Mahal ----------------
RUNES = [  # glifi in stile cirth: solo tratti dritti, box ~ 36x64 centrato
    "M0,-32 V32 M0,-32 L16,-14",
    "M0,-32 V32 M0,-10 L16,-28 M0,10 L16,-8",
    "M-10,-32 V32 M-10,-32 L12,-10 L-10,12",
    "M0,-32 V32 M-16,-14 L0,2 L16,-14",
    "M-10,-32 V32 M10,-32 V32 M-10,-4 L10,-24",
    "M0,-32 V32 M0,-32 L-16,-14 M0,-32 L16,-14",
    "M-12,-32 L12,32 M12,-32 L-12,32",
    "M0,-32 V32 M0,0 L16,18 M0,0 L-16,18",
]

def icon6():
    cx, cy = 512, 512
    s = f'    <circle cx="{cx}" cy="{cy}" r="430" fill="url(#halo)"/>\n'
    s += f'    <ellipse cx="{cx}" cy="{cy+380}" rx="260" ry="16" fill="#000" opacity="0.3"/>\n'
    # bordo dentellato
    N = 72; edge = []
    for i in range(2*N):
        a = math.pi*i/N
        rr = 366 if i % 2 == 0 else 352
        edge.append((cx + rr*math.cos(a), cy + rr*math.sin(a)))
    s += f'    <polygon points="{pts(edge)}" fill="url(#rim)" stroke="#1F2329" stroke-width="3" stroke-linejoin="round"/>\n'
    s += f'    <circle cx="{cx}" cy="{cy}" r="334" fill="url(#coin)"/>\n'
    s += f'    <circle cx="{cx}" cy="{cy}" r="334" fill="none" stroke="#FFFFFF" stroke-opacity="0.35" stroke-width="3"/>\n'
    s += f'    <circle cx="{cx}" cy="{cy}" r="316" fill="none" stroke="{G[5]}" stroke-width="3"/>\n'
    s += f'    <circle cx="{cx}" cy="{cy}" r="228" fill="none" stroke="{G[5]}" stroke-width="3"/>\n'
    # anello di rune con 8 rombi
    slots = 32; rr = 272
    for i in range(slots):
        a = 2*math.pi*i/slots - math.pi/2
        x, y = cx + rr*math.cos(a), cy + rr*math.sin(a)
        deg = math.degrees(a) + 90
        if i % 4 == 0:
            col = "url(#accent)" if i == 0 else G[1]
            s += f'    <rect x="{f(x-13)}" y="{f(y-13)}" width="26" height="26" transform="rotate(45 {f(x)} {f(y)})" fill="{col}" stroke="#1F2329" stroke-width="2"/>\n'
        else:
            d = RUNES[(i*5) % len(RUNES)]
            s += f'    <path d="{d}" transform="translate({f(x)} {f(y)}) rotate({f(deg)}) scale(0.62)" stroke="{G[5]}" stroke-width="7" stroke-linecap="square" fill="none"/>\n'
            s += f'    <path d="{d}" transform="translate({f(x+1.5)} {f(y+1.5)}) rotate({f(deg)}) scale(0.62)" stroke="#FFFFFF" stroke-opacity="0.35" stroke-width="2.5" stroke-linecap="square" fill="none"/>\n'
    # campo centrale incassato
    s += f'    <circle cx="{cx}" cy="{cy}" r="222" fill="url(#field)"/>\n'
    # martello e scalpello incrociati
    def tool(poly_list, deg):
        a = math.radians(deg); out = ""
        for poly, fill in poly_list:
            P = [(cx + p[0]*math.cos(a) - p[1]*math.sin(a), cy + p[0]*math.sin(a) + p[1]*math.cos(a)) for p in poly]
            out += f'    <polygon points="{pts(P)}" fill="{fill}" stroke="#1F2329" stroke-width="3.5" stroke-linejoin="round"/>\n'
        return out
    chisel = [
        ([(-12, -170), (12, -170), (14, 90), (-14, 90)], G[3]),
        ([(-18, -186), (18, -186), (18, -166), (-18, -166)], G[4]),
        ([(-14, 90), (14, 90), (22, 130), (0, 172), (-22, 130)], G[1]),
        ([(0, 172), (22, 130), (0, 120)], B_LIGHT),
    ]
    hammer = [
        ([(-14, -60), (14, -60), (18, 180), (-18, 180)], G[4]),
        ([(-96, -150), (96, -150), (110, -128), (110, -72), (96, -52), (-96, -52), (-110, -72), (-110, -128)], "url(#steel)"),
        ([(-96, -150), (96, -150), (84, -134), (-84, -134)], G[1]),
        ([(-26, -152), (26, -152), (26, -50), (-26, -50)], G[4]),
    ]
    s += tool(chisel, 45)
    s += tool(hammer, -45)
    s += sparkle(cx - 190, cy - 250, 26, "#FFFFFF")
    defs = ACCENT + f'''    <radialGradient id="coin" cx="38%" cy="32%" r="80%">
      <stop offset="0" stop-color="{G[2]}"/><stop offset="1" stop-color="{G[4]}"/>
    </radialGradient>
    <radialGradient id="rim" cx="38%" cy="32%" r="80%">
      <stop offset="0" stop-color="{G[3]}"/><stop offset="1" stop-color="{G[5]}"/>
    </radialGradient>
    <radialGradient id="field" cx="62%" cy="66%" r="80%">
      <stop offset="0" stop-color="{G[3]}"/><stop offset="1" stop-color="{G[5]}"/>
    </radialGradient>
'''
    return frame(s, defs)

# ---------------- 8. Ingranaggio con Gemma ----------------
def icon8():
    cx, cy = 512, 512
    s = f'    <circle cx="{cx}" cy="{cy}" r="440" fill="url(#halo)"/>\n'
    s += f'    <ellipse cx="{cx}" cy="{cy+390}" rx="250" ry="16" fill="#000" opacity="0.3"/>\n'
    Ro, Rr = 372, 300
    # 8 denti squadrati, centrati sui lati dell'ottagono della gemma
    for k in range(8):
        a = math.radians(45*k - 90)
        hw_tip, hw_root = math.radians(9.5), math.radians(14)
        tooth = [(cx + Rr*math.cos(a-hw_root), cy + Rr*math.sin(a-hw_root)),
                 (cx + Ro*math.cos(a-hw_tip), cy + Ro*math.sin(a-hw_tip)),
                 (cx + Ro*math.cos(a+hw_tip), cy + Ro*math.sin(a+hw_tip)),
                 (cx + Rr*math.cos(a+hw_root), cy + Rr*math.sin(a+hw_root))]
        s += f'    <polygon points="{pts(tooth)}" fill="{shade(math.cos(a), math.sin(a), 0.55)}" stroke="#1F2329" stroke-width="4" stroke-linejoin="round"/>\n'
        s += f'    <line x1="{f(tooth[1][0])}" y1="{f(tooth[1][1])}" x2="{f(tooth[2][0])}" y2="{f(tooth[2][1])}" stroke="#FFFFFF" stroke-opacity="0.5" stroke-width="3"/>\n'
    s += f'    <circle cx="{cx}" cy="{cy}" r="{Rr+2}" fill="url(#body)" stroke="#1F2329" stroke-width="4"/>\n'
    s += f'    <circle cx="{cx}" cy="{cy}" r="{Rr-6}" fill="none" stroke="#FFFFFF" stroke-opacity="0.3" stroke-width="3"/>\n'
    # incisioni concentriche e 8 fori di alleggerimento
    s += f'    <circle cx="{cx}" cy="{cy}" r="262" fill="none" stroke="{G[5]}" stroke-width="3"/>\n'
    for k in range(8):
        a = math.radians(45*k - 90 + 22.5)
        x, y = cx + 244*math.cos(a), cy + 244*math.sin(a)
        s += f'    <circle cx="{f(x)}" cy="{f(y)}" r="11" fill="{G[5]}" stroke="#1F2329" stroke-width="2"/>\n'
        s += f'    <circle cx="{f(x-2)}" cy="{f(y-2)}" r="4" fill="{G[2]}"/>\n'
    s += f'    <circle cx="{cx}" cy="{cy}" r="226" fill="none" stroke="{G[5]}" stroke-width="3"/>\n'
    # castone ottagonale
    Rs = 210
    sock = [(cx + Rs*math.cos(math.radians(22.5 + 45*i - 90)), cy + Rs*math.sin(math.radians(22.5 + 45*i - 90))) for i in range(8)]
    s += f'    <polygon points="{pts(sock)}" fill="#1F2329"/>\n'
    s += f'    <polygon points="{pts(sock)}" fill="none" stroke="{G[3]}" stroke-width="4" stroke-linejoin="round"/>\n'
    s += gem(cx, cy, 186, "g")
    s += sparkle(cx - 80, cy - 178, 30, "#FFFFFF")
    defs = ACCENT + f'''    <radialGradient id="body" cx="38%" cy="32%" r="85%">
      <stop offset="0" stop-color="{G[2]}"/><stop offset="1" stop-color="{G[5]}"/>
    </radialGradient>
'''
    return frame(s, defs)

if __name__ == "__main__":
    for name, fn in [("aule-1-incudine.svg", icon1), ("aule-3-martello-scintilla.svg", icon3),
                     ("aule-5-fucina.svg", icon5), ("aule-6-sigillo.svg", icon6), ("aule-8-ingranaggio.svg", icon8)]:
        open(os.path.join(OUT, name), "w").write(fn())
    print("ok")
