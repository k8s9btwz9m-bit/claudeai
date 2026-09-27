import math, os
OUT = os.path.dirname(os.path.abspath(__file__))

# Palette: grigi + blu chiaro
BG1, BG2 = "#2E333B", "#1A1D22"
G = ["#F4F6F9", "#E2E6EC", "#C8CED7", "#A7AFBB", "#838C99", "#5F6773"]
B_PALE, B_LIGHT, B_MID = "#DCF0FD", "#A9D8F6", "#6DB6E4"

def f(v): return f"{v:.1f}".rstrip("0").rstrip(".")
def pts(p): return " ".join(f"{f(x)},{f(y)}" for x, y in p)

def shade(nx, ny, nz=0.0):
    # luce da alto-sinistra, restituisce indice 0..5 nella scala di grigi
    L = (-0.55, -0.65, 0.52)
    n = math.sqrt(nx*nx + ny*ny + nz*nz) or 1
    d = (nx*L[0] + ny*L[1] + nz*L[2]) / n
    t = (d + 1) / 2
    return G[max(0, min(5, int(round((1 - t) * 5))))]

def frame(inner, defs=""):
    return f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1024 1024" width="1024" height="1024">
  <defs>
    <radialGradient id="bg" cx="50%" cy="38%" r="75%">
      <stop offset="0" stop-color="{BG1}"/>
      <stop offset="1" stop-color="{BG2}"/>
    </radialGradient>
    <radialGradient id="halo" cx="50%" cy="50%" r="50%">
      <stop offset="0" stop-color="{B_LIGHT}" stop-opacity="0.28"/>
      <stop offset="0.6" stop-color="{B_LIGHT}" stop-opacity="0.07"/>
      <stop offset="1" stop-color="{B_LIGHT}" stop-opacity="0"/>
    </radialGradient>
    <clipPath id="squircle"><rect width="1024" height="1024" rx="228"/></clipPath>
{defs}  </defs>
  <g clip-path="url(#squircle)">
    <rect width="1024" height="1024" fill="url(#bg)"/>
    <rect x="1" y="1" width="1022" height="1022" rx="227" fill="none" stroke="#FFFFFF" stroke-opacity="0.06" stroke-width="2"/>
{inner}  </g>
</svg>
'''

def sparkle(x, y, r, color=B_PALE, op=0.95):
    s = r * 0.18
    return (f'    <path d="M{f(x)},{f(y-r)} L{f(x+s)},{f(y-s)} L{f(x+r)},{f(y)} L{f(x+s)},{f(y+s)} '
            f'L{f(x)},{f(y+r)} L{f(x-s)},{f(y+s)} L{f(x-r)},{f(y)} L{f(x-s)},{f(y-s)} Z" fill="{color}" opacity="{op}"/>\n')

# ---------------- 3. Taglio Nanico ----------------
def icon3():
    cx, cy, R, r = 512, 520, 318, 172
    ang = [math.radians(22.5 + 45*i - 90) for i in range(8)]
    O = [(cx + R*math.cos(a), cy + R*math.sin(a)) for a in ang]
    I = [(cx + r*math.cos(a), cy + r*math.sin(a)) for a in ang]
    s = ""
    s += f'    <circle cx="{cx}" cy="{cy}" r="420" fill="url(#halo)"/>\n'
    # ombra di appoggio
    s += f'    <ellipse cx="{cx}" cy="{cy+R+34}" rx="220" ry="18" fill="#000" opacity="0.28"/>\n'
    facets = []
    for i in range(8):
        j = (i + 1) % 8
        M = ((O[i][0]+O[j][0])/2, (O[i][1]+O[j][1])/2)
        am = (ang[i] + ang[j]) / 2
        nx, ny = math.cos(am), math.sin(am)
        # faccetta a stella (tavola -> punto medio) : più inclinata verso il centro
        facets.append(([I[i], I[j], M], nx*0.55, ny*0.55, 0.85, i))
        # faccette di cintura
        a1 = ang[i] + math.radians(11)
        a2 = ang[j] - math.radians(11)
        facets.append(([O[i], M, I[i]], math.cos(a1), math.sin(a1), 0.45, i))
        facets.append(([M, O[j], I[j]], math.cos(a2), math.sin(a2), 0.45, i))
    for k, (poly, nx, ny, nz, i) in enumerate(facets):
        col = shade(nx, ny, nz)
        s += f'    <polygon points="{pts(poly)}" fill="{col}"/>\n'
    # faccetta accento blu (basso-destra)
    i = 1
    j = 2
    M = ((O[i][0]+O[j][0])/2, (O[i][1]+O[j][1])/2)
    s += f'    <polygon points="{pts([I[i], I[j], M])}" fill="url(#accent)"/>\n'
    s += f'    <polygon points="{pts([M, O[j], I[j]])}" fill="{B_LIGHT}" opacity="0.55"/>\n'
    # linee di taglio
    lines = ""
    for i in range(8):
        j = (i + 1) % 8
        M = ((O[i][0]+O[j][0])/2, (O[i][1]+O[j][1])/2)
        lines += f"M{f(O[i][0])},{f(O[i][1])} L{f(I[i][0])},{f(I[i][1])} "
        lines += f"M{f(M[0])},{f(M[1])} L{f(I[i][0])},{f(I[i][1])} M{f(M[0])},{f(M[1])} L{f(I[j][0])},{f(I[j][1])} "
    s += f'    <path d="{lines}" stroke="#FFFFFF" stroke-opacity="0.55" stroke-width="2" fill="none" stroke-linejoin="round"/>\n'
    s += f'    <polygon points="{pts(O)}" fill="none" stroke="{G[5]}" stroke-width="4" stroke-linejoin="round"/>\n'
    # tavola con griglia 3x3
    s += f'    <polygon points="{pts(I)}" fill="url(#table)"/>\n'
    grid = ""
    for k in (-1, 1):
        x = cx + k * r / 3 * 1.1
        grid += f"M{f(x)},{f(cy - r)} V{f(cy + r)} "
        y = cy + k * r / 3 * 1.1
        grid += f"M{f(cx - r)},{f(y)} H{f(cx + r)} "
    s += f'    <path d="{grid}" stroke="{G[3]}" stroke-width="3" clip-path="url(#tableclip)" stroke-linecap="round"/>\n'
    # cella evidenziata (in alto a destra)
    c = r / 3 * 1.1
    s += f'    <rect x="{f(cx + c + 5)}" y="{f(cy - r)}" width="{f(r)}" height="{f(r - c - 5)}" fill="{B_LIGHT}" opacity="0.6" clip-path="url(#tableclip)"/>\n'
    s += f'    <polygon points="{pts(I)}" fill="none" stroke="#FFFFFF" stroke-width="3" stroke-opacity="0.9" stroke-linejoin="round"/>\n'
    s += sparkle(O[7][0] + 6, O[7][1] - 4, 34, "#FFFFFF")
    s += sparkle(O[3][0] + 70, O[3][1] + 10, 16, B_PALE, 0.8)
    defs = f'''    <linearGradient id="table" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#FFFFFF"/><stop offset="1" stop-color="{G[1]}"/>
    </linearGradient>
    <linearGradient id="accent" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="{B_PALE}"/><stop offset="1" stop-color="{B_MID}"/>
    </linearGradient>
    <clipPath id="tableclip"><polygon points="{pts(I)}"/></clipPath>
'''
    return frame(s, defs)

# ---------------- 9. Monogramma A ----------------
def line_isect(p1, p2, p3, p4):
    x1,y1=p1;x2,y2=p2;x3,y3=p3;x4,y4=p4
    d=(x1-x2)*(y3-y4)-(y1-y2)*(x3-x4)
    a=x1*y2-y1*x2; b=x3*y4-y3*x4
    return ((a*(x3-x4)-(x1-x2)*b)/d, (a*(y3-y4)-(y1-y2)*b)/d)

def inset(poly, d):
    # poligono in senso orario (y verso il basso) -> offset verso l'interno
    n=len(poly); lines=[]
    area = sum(poly[i][0]*poly[(i+1)%n][1]-poly[(i+1)%n][0]*poly[i][1] for i in range(n))
    sgn = 1 if area > 0 else -1
    for i in range(n):
        a=poly[i]; b=poly[(i+1)%n]
        dx,dy=b[0]-a[0],b[1]-a[1]; L=math.hypot(dx,dy)
        nx,ny = -dy/L*sgn, dx/L*sgn
        lines.append(((a[0]+nx*d,a[1]+ny*d),(b[0]+nx*d,b[1]+ny*d)))
    return [line_isect(*lines[i-1], *lines[i]) for i in range(n)]

def icon9():
    A = [(446,188),(578,188),(806,832),(672,832),(512,380),(352,832),(218,832)]
    Ai = inset(A, 16)
    s = f'    <circle cx="512" cy="560" r="400" fill="url(#halo)" opacity="0.7"/>\n'
    s += f'    <ellipse cx="512" cy="858" rx="330" ry="16" fill="#000" opacity="0.3"/>\n'
    # smusso: ogni lato è un trapezio A[i],A[i+1],Ai[i+1],Ai[i] ombreggiato secondo la normale
    n = len(A)
    area = sum(A[i][0]*A[(i+1)%n][1]-A[(i+1)%n][0]*A[i][1] for i in range(n))
    sgn = 1 if area > 0 else -1
    for i in range(n):
        j=(i+1)%n
        dx,dy=A[j][0]-A[i][0],A[j][1]-A[i][1]; L=math.hypot(dx,dy)
        nx,ny = dy/L*sgn, -dx/L*sgn  # normale verso l'esterno
        s += f'    <polygon points="{pts([A[i],A[j],Ai[j],Ai[i]])}" fill="{shade(nx,ny,0.35)}"/>\n'
    s += f'    <polygon points="{pts(Ai)}" fill="url(#stone)"/>\n'
    # venature/righe nel controforma triangolare (sopra la gemma)
    rows = ""
    for y in (462, 500):
        # larghezza del vuoto interno a quella quota
        t = (y-380)/(832-380)
        xl = 512 - (512-352)*t + 14; xr = 512 + (672-512)*t - 14
        rows += f'M{f(xl+8)},{y} H{f(xr-8)} '
    s += f'    <path d="{rows}" stroke="{B_LIGHT}" stroke-width="7" stroke-linecap="round" opacity="0.85"/>\n'
    # castone
    cy = 598
    set_pts = [(372,cy),(410,cy-58),(614,cy-58),(652,cy),(614,cy+58),(410,cy+58)]
    s += f'    <polygon points="{pts(set_pts)}" fill="{G[5]}" stroke="#1F2329" stroke-width="4" stroke-linejoin="round"/>\n'
    for p in set_pts:
        s += f'    <circle cx="{f(p[0])}" cy="{f(p[1])}" r="7" fill="{G[3]}"/>\n'
    # gemma a esagono allungato con tavola
    gem = [(392,cy),(424,cy-44),(600,cy-44),(632,cy),(600,cy+44),(424,cy+44)]
    tab = [(440,cy),(456,cy-22),(568,cy-22),(584,cy),(568,cy+22),(456,cy+22)]
    norms = [(-0.6,-0.8),(0,-1),(0.6,-0.8),(0.6,0.8),(0,1),(-0.6,0.8)]
    for i in range(6):
        j=(i+1)%6
        nx,ny = norms[i]
        s += f'    <polygon points="{pts([gem[i],gem[j],tab[j],tab[i]])}" fill="{shade(nx,ny,0.5)}"/>\n'
    s += f'    <polygon points="{pts([gem[3],gem[4],tab[4],tab[3]])}" fill="url(#accent)"/>\n'
    s += f'    <polygon points="{pts(tab)}" fill="#FFFFFF"/>\n'
    s += f'    <path d="M496,{cy-22} V{cy+22} M528,{cy-22} V{cy+22} M448,{cy} H576" stroke="{G[2]}" stroke-width="2.5"/>\n'
    s += f'    <polygon points="{pts(gem)}" fill="none" stroke="#FFFFFF" stroke-opacity="0.7" stroke-width="2" stroke-linejoin="round"/>\n'
    s += sparkle(424, cy-44, 26, "#FFFFFF")
    defs = f'''    <linearGradient id="stone" x1="0" y1="0" x2="0.35" y2="1">
      <stop offset="0" stop-color="{G[2]}"/><stop offset="1" stop-color="{G[4]}"/>
    </linearGradient>
    <linearGradient id="accent" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="{B_PALE}"/><stop offset="1" stop-color="{B_MID}"/>
    </linearGradient>
'''
    return frame(s, defs)

# ---------------- 1. Gemma-Griglia ----------------
def icon1():
    cx, cy, R = 512, 512, 300
    tilt = math.radians(-22); rot = math.radians(14)
    def P(lat, lon):
        x = math.cos(lat)*math.sin(lon+rot); y = -math.sin(lat); z = math.cos(lat)*math.cos(lon+rot)
        # inclinazione attorno all'asse X
        y2 = y*math.cos(tilt) - z*math.sin(tilt); z2 = y*math.sin(tilt) + z*math.cos(tilt)
        return (x, y2, z2)
    NLAT, NLON = 8, 14
    lats = [math.radians(-90 + 180*i/NLAT) for i in range(NLAT+1)]
    lons = [2*math.pi*j/NLON for j in range(NLON)]
    highlight = {(4,1):"a", (3,2):"b", (5,3):"b", (2,12):"b"}
    cells = []
    for i in range(NLAT):
        for j in range(NLON):
            q = [P(lats[i],lons[j]), P(lats[i],lons[(j+1)%NLON]), P(lats[i+1],lons[(j+1)%NLON]), P(lats[i+1],lons[j])]
            # normale = centroide (faccetta piana di una sfera)
            c = [sum(v[k] for v in q)/4 for k in range(3)]
            if c[2] <= 0.02: continue
            poly = [(cx+R*v[0], cy+R*v[1]) for v in q]
            # rimuovi vertici duplicati (poli)
            uniq=[]
            for p in poly:
                if not uniq or math.hypot(p[0]-uniq[-1][0],p[1]-uniq[-1][1])>0.5: uniq.append(p)
            if len(uniq)>2 and math.hypot(uniq[0][0]-uniq[-1][0],uniq[0][1]-uniq[-1][1])<0.5: uniq.pop()
            cells.append((c[2], uniq, c, highlight.get((i,j))))
    cells.sort(key=lambda t: t[0])
    s = f'    <circle cx="{cx}" cy="{cy}" r="430" fill="url(#halo)"/>\n'
    s += f'    <ellipse cx="{cx}" cy="{cy+R+40}" rx="210" ry="18" fill="#000" opacity="0.3"/>\n'
    s += f'    <circle cx="{cx}" cy="{cy}" r="{R+2}" fill="{G[5]}"/>\n'
    for z, poly, c, hl in cells:
        if hl == "a": fill = "url(#accent)"
        elif hl == "b": fill = B_LIGHT
        else:
            L=(-0.5,-0.62,0.6); n=math.sqrt(sum(v*v for v in c))
            d=(c[0]*L[0]+c[1]*L[1]+c[2]*L[2])/n
            fill = G[max(0,min(5,int(round((1-d)*6.2))))]
        s += f'    <polygon points="{pts(poly)}" fill="{fill}" stroke="#FFFFFF" stroke-opacity="0.55" stroke-width="2" stroke-linejoin="round"/>\n'
    s += f'    <circle cx="{cx}" cy="{cy}" r="{R}" fill="url(#pearl)"/>\n'
    s += f'    <circle cx="{cx}" cy="{cy}" r="{R}" fill="none" stroke="{G[5]}" stroke-width="4"/>\n'
    s += sparkle(cx-150, cy-205, 40, "#FFFFFF")
    s += sparkle(cx+235, cy-150, 16, B_PALE, 0.85)
    defs = f'''    <radialGradient id="pearl" cx="36%" cy="30%" r="75%">
      <stop offset="0" stop-color="#FFFFFF" stop-opacity="0.45"/>
      <stop offset="0.45" stop-color="#FFFFFF" stop-opacity="0"/>
      <stop offset="1" stop-color="{BG2}" stop-opacity="0.25"/>
    </radialGradient>
    <linearGradient id="accent" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="{B_PALE}"/><stop offset="1" stop-color="{B_MID}"/>
    </linearGradient>
'''
    return frame(s, defs)

if __name__ == "__main__":
  for name, fn in [("arkenstone-1-gemma-griglia.svg", icon1), ("arkenstone-3-taglio-nanico.svg", icon3), ("arkenstone-9-monogramma-a.svg", icon9)]:
    open(os.path.join(OUT, name), "w").write(fn())
  print("ok")
